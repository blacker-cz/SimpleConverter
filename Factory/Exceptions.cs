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
}
