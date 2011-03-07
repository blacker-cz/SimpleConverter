using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel.Composition;
using SimpleConverter.Contract;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    [Export(typeof(IPlugin))]
    [PluginMetadata("Beamer 2 Powerpoint", "0.0.0.1", "Plugin for conversion from Beamer to PowerPoint")]
    public class Connector : IPlugin, IMessenger
    {
        private System.Windows.FrameworkElement _visual;

        #region IPlugin implementation

        public event SendMessageDelegate SendMessageEvent;

        public event ProgressDelegate ProgressEvent;

        public void ConvertDocument(string filename, string outputDirectory = "")
        {
            // register Messenger
            Messenger.Instance.Add(this);

            // hardcoded for debugging purposes - will use threads
            Parser parser;

            System.IO.FileStream reader;
            reader = new System.IO.FileStream(filename, System.IO.FileMode.Open);
            Scanner scanner = new Scanner(reader);

            parser = new Parser(scanner);

            Messenger.Instance.SendMessage("Started parsing.");

            bool ok = parser.Parse();

            reader.Close();

            #region Debug serialization
#if DEBUG
            if (ok)
            {
                System.Xml.Serialization.XmlSerializer x = new System.Xml.Serialization.XmlSerializer(parser.Document.GetType());
                System.IO.StreamWriter writer = new System.IO.StreamWriter(@"document.xml");
                x.Serialize(writer, parser.Document);

                Messenger.Instance.SendMessage(@"Document tree serialized to ""document.xml""");
            }
#endif
            #endregion // Debug serialization

            Messenger.Instance.SendMessage("Parsing done!");
        }

        public System.Windows.FrameworkElement GetVisual()
        {
            if (_visual == null)    // keep plugin view instantiated
                _visual = new SettingsView();

            return _visual;
        }

        public bool ValidateFile(string filename)
        {
            // todo: refactor this!!
            if (System.IO.Path.GetExtension(filename) == ".tex")
                return true;
            return false;
        }

        #endregion

        #region IMessenger implementation
        /// <summary>
        /// Fire event with message
        /// todo: thread safe?
        /// </summary>
        /// <param name="message">Text message</param>
        /// <param name="level">Message level</param>
        public void SendMessage(string message, MessageLevel level = MessageLevel.INFO)
        {
            try
            {
                if (SendMessageEvent != null)
                {
                    SendMessageEvent(message, level);
                }
            }
            catch
            {
                
            }
        }
        #endregion
    }
}
