using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SimpleConverter
{
    /// <summary>
    /// Message wrapper class
    /// </summary>
    public class ListMessage
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="message">Message text</param>
        /// <param name="level">Message level</param>
        public ListMessage(string message, Contract.MessageLevel level = Contract.MessageLevel.INFO)
        {
            Message = message;
            switch (level)
            {
                case SimpleConverter.Contract.MessageLevel.INFO:
                    Icon = (System.Windows.Media.ImageSource)System.Windows.Application.Current.FindResource("iconInfo");
                    break;
                case SimpleConverter.Contract.MessageLevel.WARNING:
                    Icon = (System.Windows.Media.ImageSource)System.Windows.Application.Current.FindResource("iconWarning");
                    break;
                case SimpleConverter.Contract.MessageLevel.ERROR:
                    Icon = (System.Windows.Media.ImageSource)System.Windows.Application.Current.FindResource("iconError");
                    break;
                default:
                    throw new ArgumentException("Unknown message level.");
            }
        }

        /// <summary>
        /// Message text
        /// </summary>
        public string Message { get; private set; }

        /// <summary>
        /// Message icon
        /// </summary>
        public System.Windows.Media.ImageSource Icon { get; private set; }
    }
}
