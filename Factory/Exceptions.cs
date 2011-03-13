using System;
using System.Runtime.Serialization;

namespace SimpleConverter.Factory
{
    /// <summary>
    /// Custom exception used for plugin handling errors
    /// </summary>
    [Serializable]
    public class PluginLoaderException : Exception
    {
        public PluginLoaderException(string message)
            : base(message) { }

        public PluginLoaderException(string message, Exception innerException)
            : base(message, innerException) { }
        
        protected PluginLoaderException(SerializationInfo info, StreamingContext context)
            : base(info, context) { }
    }

    /// <summary>
    /// Custom exception for invalid state error
    /// </summary>
    [Serializable]
    public class InvalidStateException : Exception
    {
        public InvalidStateException()
            : base(@"Invalid application state.") { }

        public InvalidStateException(string message)
            : base(message) { }

        public InvalidStateException(string message, Exception innerException)
            : base(message, innerException) { }

        protected InvalidStateException(SerializationInfo info, StreamingContext context)
            : base(info, context) { }
    }

    /// <summary>
    /// Custom exception for invalid argument error
    /// </summary>
    [Serializable]
    public class InvalidArgumentException : Exception
    {
        public InvalidArgumentException()
            : base(@"Invalid argument passed.") { }

        public InvalidArgumentException(string message)
            : base(message) { }

        public InvalidArgumentException(string message, Exception innerException)
            : base(message, innerException) { }

        protected InvalidArgumentException(SerializationInfo info, StreamingContext context)
            : base(info, context) { }
    }
}
