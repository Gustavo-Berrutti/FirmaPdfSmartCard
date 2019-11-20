using iText.Kernel.Pdf;
using iText.Signatures;

using Net.Pkcs11Interop.Common;
using Net.Pkcs11Interop.HighLevelAPI;
using Net.Pkcs11Interop.HighLevelAPI.Factories;
using Org.BouncyCastle.X509;
using Serilog;
using Signer.Impl;
using Signer.Pdf;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Threading.Tasks;

namespace Signer.Pkcs11Interop
{
    class PkcsInteropSigner : PdfSignerBase
    {
        ILogger logger = Log.ForContext<PkcsInteropSigner>();

        IExternalSignature externalSignature;

        IPkcs11Library pkcs11Library;
        ISlot slot;
        ISession session;
        Org.BouncyCastle.X509.X509Certificate[] certPath;
        IObjectHandle privatekeyHandle;

        public override event EventHandler<SignedFileEventArgs> SignedFileEvent;

        private static string pkcs11LibraryPath;


        static PkcsInteropSigner()
        {
            pkcs11LibraryPath = ConfigurationManager.AppSettings.Get("PkcsInteropSigner.dllPath");
        }

        public PkcsInteropSigner()
        {
            Pkcs11InteropFactories factories = new Pkcs11InteropFactories();
            pkcs11Library = factories.Pkcs11LibraryFactory.LoadPkcs11Library(factories, pkcs11LibraryPath, AppType.MultiThreaded);

            var slots = pkcs11Library.GetSlotList(SlotsType.WithTokenPresent);
            if (slots.Count == 0)
            {
                throw new InstanceCreationException("Por favor, verifique que el lector está conectado y que la cédula fue insertada correctamente.");

            }
            slot = slots[0];
        }

        ~PkcsInteropSigner()
        {
            if (pkcs11Library != null) {
                pkcs11Library.Dispose();
                pkcs11Library.Dispose();
            }
        
        }

        public override string Pin
        {
            set
            {
                try
                {
                    if (session != null)
                        session.Dispose();

                    pin = value;
                    session = slot.OpenSession(SessionType.ReadOnly);

                    session.Login(CKU.CKU_USER, pin);
                    LoadCertPrivateKeyHandle();
                    externalSignature = new Pkcs11ExternalSignature(session, privatekeyHandle);
                }
                catch (Pkcs11Exception pkc11Exc)
                {
                    if (pkc11Exc.RV == CKR.CKR_PIN_INCORRECT || pkc11Exc.RV == CKR.CKR_PIN_INVALID || pkc11Exc.RV == CKR.CKR_PIN_LOCKED)
                    {
                        logger.Error("Pin incorrecto, error devuelto: {pinErrorValue}", pkc11Exc.RV);
                        throw new InstanceCreationException($"El pin ingresado es incorrecto, inválido o está bloqueado ({pkc11Exc.RV}).");
                    }
                    else
                    {
                        logger.Error($"Error inesperado usando el pin.\n{pkc11Exc.Message} ({pkc11Exc.RV})");
                        throw new InstanceCreationException("Ha ocurrido un error inesperado.\n" + pkc11Exc.Message);
                    }
                }

            }
        }


        public override void Dispose()
        {
            if (session != null)
                session.Dispose();
            if (pkcs11Library != null)
                pkcs11Library.Dispose();

        }


        public override bool RequiresPin => true;
        

        public override async Task SignPdf(FileInfo fileInfo)
        {
            await SignPdf(fileInfo.FullName);
        }

        public override async Task SignPdf(string filePath)
        {
            await SignPdf(filePath, File.ReadAllBytes(filePath));
        }

        public override async Task SignPdf(string fileId, byte[] fileData)
        {
            SignedFileEventArgs eventArgs = null;

            if (String.IsNullOrWhiteSpace(pin) || privatekeyHandle == null)
                eventArgs = new SignedFileEventArgs(null, fileId, true, "El pin es requerido.\nO bien no fue ingresado, o era incorrecto.");

            else
            {
                using (var memStream = new MemoryStream(fileData))
                using (PdfReader reader = new PdfReader(memStream))
                {
                    StampingProperties properties = new StampingProperties();
                    properties.UseAppendMode();

                    using (MemoryStream outputStream = new MemoryStream())
                    {
                        try
                        {
                            PdfSigner signer = new PdfSigner(reader, outputStream, properties);
                            /* signer.SetCertificationLevel(certificationLevel);
                             PdfSignatureAppearance appearance = signer.GetSignatureAppearance().SetReason(reason).SetLocation(location
                                 ).SetReuseAppearance(setReuseAppearance);

                             signer.SetFieldName(name);*/
                            // Creating the signature

                            signer.SignDetached(externalSignature, certPath, null, null, null, 0, PdfSigner.CryptoStandard.CADES);
                            eventArgs = new SignedFileEventArgs(outputStream.ToArray(), fileId, false, "");
                        }
                        catch (Exception e)
                        {
                            logger.Error(e, "Error firmando archivo {fileId}\n{mensaje}", fileId, e.Message);
                            eventArgs = new SignedFileEventArgs(null, fileId, true, e.Message);
                        }
                    }
                }
            }
            OnRaiseSignedFileEvent(eventArgs);
            await Task.CompletedTask;
        }

        private void OnRaiseSignedFileEvent(SignedFileEventArgs e)
        {
            EventHandler<SignedFileEventArgs> handler = SignedFileEvent;

            if (handler != null)
            {
                handler(this, e);
            }
        }


        private void LoadCertPrivateKeyHandle()
        {

            //Carga del certificado
            List<IObjectAttribute> template = new List<IObjectAttribute>();
            var f = new ObjectAttributeFactory();
            template.Add(f.Create(CKA.CKA_CLASS, CKO.CKO_CERTIFICATE));
            template.Add(f.Create(CKA.CKA_CERTIFICATE_TYPE, CKC.CKC_X_509));

            List<IObjectHandle> pubKeyObjectHandle = session.FindAllObjects(template);

            List<CKA> pubKeyAttrsToRead = new List<CKA>();
            pubKeyAttrsToRead.Add(CKA.CKA_LABEL);
            pubKeyAttrsToRead.Add(CKA.CKA_ID);
            pubKeyAttrsToRead.Add(CKA.CKA_VALUE);

            //La cédula tiene sólo un certificado
            List<IObjectAttribute> pubKeyAttributes = session.GetAttributeValue(pubKeyObjectHandle[0], pubKeyAttrsToRead);
            var certArray = pubKeyAttributes[2].GetValueAsByteArray();
            X509CertificateParser _x509CertificateParser = new X509CertificateParser();
            Org.BouncyCastle.X509.X509Certificate bcCert = _x509CertificateParser.ReadCertificate(certArray);

            //El identificador del certificado
            string ckaLabel = pubKeyAttributes[0].GetValueAsString();

            //Recupera el handle a la clave privada utilizando el identificador del certificado
            template.Clear();
            template.Add(f.Create(CKA.CKA_CLASS, CKO.CKO_PRIVATE_KEY));
            template.Add(f.Create(CKA.CKA_KEY_TYPE, CKK.CKK_RSA));

            template.Add(f.Create(CKA.CKA_LABEL, ckaLabel));

            privatekeyHandle = session.FindAllObjects(template)[0];

            certPath = new Org.BouncyCastle.X509.X509Certificate[] { bcCert };

        }

    }
}
