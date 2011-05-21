using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel.Composition;
using SimpleConverter.Contract;
using NDesk.Options;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    [Export(typeof(IPlugin))]
    [PluginMetadata("Beamer 2 Powerpoint", "0.9.0.1", "Plugin for conversion from Beamer to PowerPoint")]
    public class Connector : IPlugin, IMessenger
    {
        /// <summary>
        /// Plugin UI
        /// </summary>
        private System.Windows.FrameworkElement _visual;

        /// <summary>
        /// PowerPoint builder class (for creation of output)
        /// </summary>
        private PowerPointBuilder _builder;

        /// <summary>
        /// Command line options parser
        /// </summary>
        private OptionSet _options;

        #region IPlugin implementation

        /// <summary>
        /// Event for sending messages
        /// </summary>
        public event SendMessageDelegate SendMessageEvent;

        /// <summary>
        /// Event for reporting current progress
        /// </summary>
        public event ProgressDelegate ProgressEvent;

        /// <summary>
        /// Constructor
        /// </summary>
        public Connector()
        {
            // register Messenger
            Messenger.Instance.Add(this);

            _options = new OptionSet() {
			    { "n|nadjust",  "don't adjust image and table size", v => Settings.Instance.AdjustSize = v == null },
			    { "e|extract",  "extract nested elements", v => Settings.Instance.NestedAsText = v == null },
            };
        }

        /// <summary>
        /// Initialize plugin before document conversion.
        /// Start PowerPoint
        /// </summary>
        public void Init()
        {
            try
            {
                _builder = new PowerPointBuilder();
                // setup progress delegate
                _builder.Progress = new ProgressDelegate(ProgressInfo);
            }
            catch (PowerPointApplicationException ex)
            {
                Messenger.Instance.SendMessage(ex.Message, MessageLevel.ERROR);
                throw ex;    // propagate exception (to end document processing loop)
            }
        }

        /// <summary>
        /// Conversion is completed, free plugin resources
        /// </summary>
        public void Done()
        {
            // close PowerPoint and opened presentations
            _builder.Close();
            _builder = null;
        }

        /// <summary>
        /// Converts document
        /// </summary>
        /// <param name="filename">Filename of input document</param>
        /// <param name="outputDirectory">Output directory for converted document</param>
        public void ConvertDocument(string filename, string outputDirectory = "")
        {
            // set progress info to initial value
            ProgressInfo(0);

            #region Analysis of Beamer document
            Messenger.Instance.SendMessage(System.IO.Path.GetFileName(filename) + " - Started parsing.");

            Parser parser;

            System.IO.FileStream reader;
            reader = new System.IO.FileStream(filename, System.IO.FileMode.Open);
            Scanner scanner = new Scanner(reader);
            scanner.Filename = System.IO.Path.GetFileName(filename);

            parser = new Parser(scanner);

            bool ok = parser.Parse();

            reader.Close();

            if (!ok)
                throw new DocumentException(System.IO.Path.GetFileName(filename) + " - Error processing input !");

            Messenger.Instance.SendMessage(System.IO.Path.GetFileName(filename) + " - Parsing done!");

            #endregion // Analysis of Beamer document

            // check if there are slides in presentation
            if (parser.SlideCount == 0)
            {
                Messenger.Instance.SendMessage(System.IO.Path.GetFileName(filename) + " - Empty presentation - output omitted.", MessageLevel.ERROR);
                return;
            }

            ProgressInfo(PowerPointBuilder.BasicProgress);

            #region Building output document

            Messenger.Instance.SendMessage(System.IO.Path.GetFileName(filename) + " - Started building output.");

            if (_builder == null)
                throw new InvalidOperationException("Plugin wasn't initialized yet");

            try
            {
                _builder.Build(filename, outputDirectory, parser.Document, parser.SlideCount, parser.SectionTable, parser.FrametitleTable);
            }
            catch (Exception ex)
            {
                if (ex is PowerPointApplicationException || ex is DocumentBuilderException)
                {
                    Messenger.Instance.SendMessage(ex.Message, MessageLevel.ERROR);

                    Messenger.Instance.SendMessage(System.IO.Path.GetFileName(filename) + " - Presentation couldn't be converted.", MessageLevel.ERROR);
                }

                throw ex;   // propagate exception
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
            if (System.IO.Path.GetExtension(filename) == ".tex" && System.IO.File.Exists(filename))
                return true;
            return false;
        }

        /// <summary>
        /// Process console options (and setup internal logic)
        /// </summary>
        /// <param name="options">List of options</param>
        /// <returns>List of not-processed parameters</returns>
        public List<string> ConsoleOptions(List<string> options)
        {
            List<string> extra;
            try
            {
                extra = _options.Parse(options);
            }
            catch (OptionException e)
            {
                return options;
            }

            return extra;
        }

        /// <summary>
        /// Print console help
        /// </summary>
        public void ConsoleHelp()
        {
            _options.WriteOptionDescriptions(System.Console.Out);
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
