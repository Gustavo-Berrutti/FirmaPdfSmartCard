using Signer.Pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Signer.Impl
{
    internal class FirmaDocumentos : IFirmaDocumentos
    {
        

        PdfSignerBase pdfSigner;

        public FirmaDocumentos(PdfSignerBase pdfSigner)
        {
            this.pdfSigner = pdfSigner;
        }

        public PdfSignerBase getPdfSigner()
        {
            return pdfSigner;
        }
    }
}
