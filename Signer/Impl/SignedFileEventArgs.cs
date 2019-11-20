using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Signer.Impl
{
    public class SignedFileEventArgs:EventArgs
    {
        
        public FileInfo File { get; private set; }

        public byte[] SignedContent { get; private set; }

        public String FileId { get; private set; }

        public Boolean HasError { get; private set; }

        public String Message { get; private set; }

        internal SignedFileEventArgs(String fileNameFullPath, byte[] fileData, bool error, String message) :this(new FileInfo(fileNameFullPath), fileData, error,message){}

        internal SignedFileEventArgs(FileInfo fileInfo, byte[] fileData,bool error, String message) : base()
        {
            File = fileInfo;
            SignedContent = fileData;
            FileId = fileInfo.FullName;
            HasError = error;
            Message = message;
        }

        internal SignedFileEventArgs( byte[] fileData,String fileId, bool error, String message) : this(new FileInfo(fileId), fileData, error, message){}


    }
}
