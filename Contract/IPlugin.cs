using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SimpleConverter.Contract
{
    /// <summary>
    /// Interface for plugins
    /// </summary>
    public interface IPlugin
    {
        /// <summary>
        /// Event for text messages from plugin
        /// </summary>
        event SendMessageDelegate SendMessageEvent;

        /// <summary>
        /// Event for progress information (from 0 to 100)
        /// </summary>
        event ProgressDelegate ProgressEvent;

        /// <summary>
        /// Converts document
        /// </summary>
        /// <param name="filename">Filename of input document</param>
        /// <param name="outputDirectory">Output directory for converted document</param>
        void ConvertDocument(string filename, string outputDirectory = "");

        /// <summary>
        /// Content of the settings tab
        /// </summary>
        /// <returns></returns>
        System.Windows.FrameworkElement GetVisual();

        /// <summary>
        /// Check if filename is of supported type
        /// </summary>
        /// <param name="filename">Filename</param>
        /// <returns>True if is supported, false otherwise</returns>
        bool ValidateFile(string filename);
    }

    /// <summary>
    /// Plugin messaging delagate
    /// </summary>
    /// <param name="message">Text of message</param>
    /// <param name="level">Level of message</param>
    public delegate void SendMessageDelegate(string message, MessageLevel level = MessageLevel.INFO);

    /// <summary>
    /// Progress information delegate
    /// </summary>
    /// <param name="progress">Progress in % (0 - 100)</param>
    public delegate void ProgressDelegate(int progress);

    /// <summary>
    /// Level of message
    /// </summary>
    public enum MessageLevel { INFO = 0, WARNING, ERROR };
}
