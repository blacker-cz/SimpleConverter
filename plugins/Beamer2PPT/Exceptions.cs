﻿using System;
using System.Runtime.Serialization;
using SimpleConverter.Contract;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    /// <summary>
    /// Custom exception used for PowerPoint manipulation errors
    /// </summary>
    [Serializable]
    class PowerPointApplicationException : InitException
    {
        public PowerPointApplicationException(string message)
            : base(message) { }

        public PowerPointApplicationException(string message, Exception innerException)
            : base(message, innerException) { }

        protected PowerPointApplicationException(SerializationInfo info, StreamingContext context)
            : base(info, context) { }
    }

    /// <summary>
    /// Custom exception used for output document building errors
    /// </summary>
    [Serializable]
    class DocumentBuilderException : DocumentException
    {
        public DocumentBuilderException(string message)
            : base(message) { }

        public DocumentBuilderException(string message, Exception innerException)
            : base(message, innerException) { }

        protected DocumentBuilderException(SerializationInfo info, StreamingContext context)
            : base(info, context) { }
    }
}
