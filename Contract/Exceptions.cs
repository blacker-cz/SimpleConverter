using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.Serialization;

namespace SimpleConverter.Contract
{
    /// <summary>
    /// Custom exception used for errors during plugin initialization
    /// </summary>
    [Serializable]
    public class InitException : Exception
    {
        public InitException(string message)
            : base(message) { }

        public InitException(string message, Exception innerException)
            : base(message, innerException) { }

        protected InitException(SerializationInfo info, StreamingContext context)
            : base(info, context) { }
    }

    /// <summary>
    /// Custom exception used for errors during document conversion
    /// </summary>
    [Serializable]
    public class DocumentException : Exception
    {
        public DocumentException(string message)
            : base(message) { }

        public DocumentException(string message, Exception innerException)
            : base(message, innerException) { }

        protected DocumentException(SerializationInfo info, StreamingContext context)
            : base(info, context) { }
    }
}
