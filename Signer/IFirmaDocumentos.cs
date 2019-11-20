using Signer.Impl;
using Signer.Pdf;
using System;
using System.IO;

namespace Signer
{
    public interface IFirmaDocumentos
    {
        PdfSignerBase getPdfSigner();
    }
}
