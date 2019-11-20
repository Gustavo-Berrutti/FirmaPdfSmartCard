using System;
using System.Runtime.Serialization;

namespace Signer
{
    [Serializable]
    public class InstanceCreationException : Exception
    {
        public InstanceCreationException()
        {
        }

        public InstanceCreationException(string message) : base(message)
        {
        }

        public InstanceCreationException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected InstanceCreationException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}