using Signer.Impl;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Signer.Pdf
{
    abstract public class PdfSignerBase:IDisposable
    {

        protected String pin;

        protected PdfSignerBase() {
            
        }

        public abstract String Pin { set ;}

        public abstract event EventHandler<SignedFileEventArgs> SignedFileEvent;

        public abstract void Dispose();

        public virtual  Boolean RequiresPin => false ;

        

        public abstract Task SignPdf(FileInfo fileInfo);
        public abstract Task SignPdf(String filePath);
        public abstract Task SignPdf(String fileId, byte[] fileData);
    }
}
