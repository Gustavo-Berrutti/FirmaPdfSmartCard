using Signer.Impl;
using Signer.IsaHub;
using Signer.Pdf;
using Signer.Pkcs11Interop;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Text;

namespace Signer
{
    public class FirmaDocumentosFactory
    {
        private static FirmaDocumentosFactory firmaDocs= new FirmaDocumentosFactory();

        private static String PdfSignerImpl = (ConfigurationManager.AppSettings.Get("PdfSigner.Impl")??"").ToLower();

        public static IFirmaDocumentos GetInstance() {
            return new FirmaDocumentos(GetPdfSignerImpl());   
        }

        private static PdfSignerBase GetPdfSignerImpl()
        {
            switch (PdfSignerImpl) {
                case "iscert":
                    return new IsaSigner();
                case "pkcsinterop":
                    return new PkcsInteropSigner();
                default:
                    return new PkcsInteropSigner();
            }
        }
    }
}
