using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SimpleConverter.Contract;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    /// <summary>
    /// Public interface for messengers
    /// </summary>
    public interface IMessenger
    {
        /// <summary>
        /// Send message
        /// </summary>
        /// <param name="message">Message text</param>
        /// <param name="level">Message level</param>
        void SendMessage(string message, MessageLevel level = MessageLevel.INFO);
    }
}
