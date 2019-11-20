using iText.Signatures;
using Net.Pkcs11Interop.Common;
using Net.Pkcs11Interop.HighLevelAPI;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Asn1.X509;
using Org.BouncyCastle.Crypto;
using Org.BouncyCastle.Crypto.Digests;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Signer.Pkcs11Interop

{
    class Pkcs11ExternalSignature:IExternalSignature
    {
        
        IObjectHandle _privateKeyHandle;
        ISession session;

        public Pkcs11ExternalSignature(ISession session, IObjectHandle privateKeyHandle)
        {
            this.session= session;
            _privateKeyHandle = privateKeyHandle;
        }

        //Se utilizan algoritmos fijos. Se puede modificar facilmente.
        public string GetEncryptionAlgorithm()
        {
            return "RSA";
        }


        public string GetHashAlgorithm()
        {

            return "SHA256"; 
        }


        public byte[] Sign(byte[] message)
        {

            byte[] digest = null;
            byte[] digestInfo = null;

            switch (GetHashAlgorithm())
            {
                case "SHA1":
                    digest = ComputeDigest(new Sha1Digest(), message);
                    digestInfo = CreateDigestInfo(digest, "1.3.14.3.2.26");
                    break;
                case "SHA256":
                    digest = ComputeDigest(new Sha256Digest(), message);
                    digestInfo = CreateDigestInfo(digest, "2.16.840.1.101.3.4.2.1");
                    break;
                case "SHA384":
                    digest = ComputeDigest(new Sha384Digest(), message);
                    digestInfo = CreateDigestInfo(digest, "2.16.840.1.101.3.4.2.2");
                    break;
                case "SHA512":
                    digest = ComputeDigest(new Sha512Digest(), message);
                    digestInfo = CreateDigestInfo(digest, "2.16.840.1.101.3.4.2.3");
                    break;
            }
            var m = new Net.Pkcs11Interop.HighLevelAPI.Factories.MechanismFactory();
            
            using (IMechanism mechanism = m.Create(CKM.CKM_RSA_PKCS))
                return session.Sign(mechanism, _privateKeyHandle, digestInfo);
        }



        private byte[] ComputeDigest(IDigest digest, byte[] data)
        {
            if (digest == null)
                throw new ArgumentNullException("digest");

            if (data == null)
                throw new ArgumentNullException("data");

            byte[] hash = new byte[digest.GetDigestSize()];

            digest.Reset();
            digest.BlockUpdate(data, 0, data.Length);
            digest.DoFinal(hash, 0);

            return hash;
        }

        private byte[] CreateDigestInfo(byte[] hash, string hashOid)
        {
            DerObjectIdentifier derObjectIdentifier = new DerObjectIdentifier(hashOid);
            AlgorithmIdentifier algorithmIdentifier = new AlgorithmIdentifier(derObjectIdentifier, null);
            DigestInfo digestInfo = new DigestInfo(algorithmIdentifier, hash);
            return digestInfo.GetDerEncoded();
        }

    }
}
