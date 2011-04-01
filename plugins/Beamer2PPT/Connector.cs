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

        /// <summary>
        /// Converts document
        /// </summary>
        /// <param name="filename">Filename of input document</param>
        /// <param name="outputDirectory">Output directory for converted document</param>
        public void ConvertDocument(string filename, string outputDirectory = "")
        {
            // register Messenger
            Messenger.Instance.Add(this);

            // set progress info to initial value
            ProgressInfo(0);

            #region Analysis of Beamer document
            Messenger.Instance.SendMessage("Started parsing.");

            Parser parser;

            System.IO.FileStream reader;
            reader = new System.IO.FileStream(filename, System.IO.FileMode.Open);
            Scanner scanner = new Scanner(reader);

            parser = new Parser(scanner);

            bool ok = parser.Parse();

            reader.Close();

            if (!ok)    // todo: print message or not?
                return;

            Messenger.Instance.SendMessage("Parsing done!");

            #endregion // Analysis of Beamer document

            #region Debug serialization
#if DEBUG
            System.Xml.Serialization.XmlSerializer x = new System.Xml.Serialization.XmlSerializer(parser.Document.GetType());
            System.IO.StreamWriter writer = new System.IO.StreamWriter(@"document.xml");
            x.Serialize(writer, parser.Document);

            Messenger.Instance.SendMessage(@"Document tree serialized to ""document.xml""");
#endif
            #endregion // Debug serialization

            // check if there are slides in presentation
            if (parser.SlideCount == 0)
            {
                Messenger.Instance.SendMessage("Empty presentation - output omitted.", MessageLevel.ERROR);
                return;
            }

            ProgressInfo(PowerPointBuilder.BasicProgress);

            #region Building output document

            Messenger.Instance.SendMessage("Started building output.");

            PowerPointBuilder builder;

            try
            {
                builder = new PowerPointBuilder(filename, outputDirectory, parser.Document, parser.SlideCount, parser.SectionTable, parser.FrametitleTable);
                // setup progress delegate
                builder.Progress = new ProgressDelegate(ProgressInfo);
            }
            catch (PowerPointApplicationException ex)
            {
                Messenger.Instance.SendMessage(ex.Message, MessageLevel.ERROR);
                throw new Exception(ex.Message, ex);    // propagate exception (to end document processing loop)
            }

            try
            {
                builder.Build();

                Messenger.Instance.SendMessage("Output building done.");
            }
            catch (DocumentBuilderException ex)
            {
                Messenger.Instance.SendMessage(ex.Message, MessageLevel.ERROR);
            }
            finally
            {
                // close PowerPoint and opened presentations
                builder.Close();
            }

            #endregion // Building output document
        }

        /// <summary>
        /// Content of the settings tab
        /// </summary>
        /// <returns></returns>
        public System.Windows.FrameworkElement GetVisual()
        {
            if (_visual == null)    // keep plugin view instantiated
                _visual = new SettingsView();

            return _visual;
        }

        /// <summary>
        /// Check if filename is of supported type
        /// </summary>
        /// <param name="filename">Filename</param>
        /// <returns>True if is supported, false otherwise</returns>
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

        /// <summary>
        /// Fire event with progress info
        /// </summary>
        /// <param name="progress">Progress information</param>
        public void ProgressInfo(int progress)
        {
            try
            {
                if (ProgressEvent != null)
                {
                    ProgressEvent(progress);
                }
            }
            catch
            {

            }
        }

    }
}
