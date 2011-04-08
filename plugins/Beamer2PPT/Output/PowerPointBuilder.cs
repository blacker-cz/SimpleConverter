using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    class PowerPointBuilder
    {
        #region Public variables and properties

        /// <summary>
        /// Magic number for progress after document parsing
        /// </summary>
        public const int BasicProgress = 20;

        /// <summary>
        /// Progress information delegate
        /// </summary>
        public Contract.ProgressDelegate Progress { get; set; }

        #endregion // Public variables and properties

        #region Private variables

        /// <summary>
        /// Current progress
        /// </summary>
        private int _currentProgress = BasicProgress;

        /// <summary>
        /// Filename of output file (without extension)
        /// </summary>
        private string _filename;

        /// <summary>
        /// Folder containing input file
        /// </summary>
        private string _inputFolder;

        /// <summary>
        /// Document tree
        /// </summary>
        private Node _document;

        /// <summary>
        /// Slide count
        /// </summary>
        private int _slideCount;

        /// <summary>
        /// PowerPoint application instance
        /// </summary>
        private PowerPoint.Application _pptApplication;

        /// <summary>
        /// PowerPoint presentation
        /// </summary>
        private PowerPoint.Presentation _pptPresentation;

        /// <summary>
        /// List of section definition records
        /// </summary>
        private List<SectionRecord> _sectionTable;

        /// <summary>
        /// Table of frame titles
        /// </summary>
        private Dictionary<int, FrametitleRecord> _frametitleTable;

        /// <summary>
        /// Table with global title page settings (author, title, date, etc.)
        /// </summary>
        private Dictionary<string, List<Node>> _titlePageSettings;

        /// <summary>
        /// Internal counter for currently processed slide
        /// </summary>
        private int _currentSlide;

        /// <summary>
        /// Internal counter for slide index (because of overlays - one slide definition can generate more then one slide)
        /// </summary>
        private int _slideIndex;

        /// <summary>
        /// Base font size
        /// </summary>
        private float _baseFontSize = 11.0f;

        #endregion // Private variables

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="filename">Filename of input file (used for output file)</param>
        /// <param name="outputPath">Output directory</param>
        /// <param name="document">Document tree</param>
        /// <param name="slideCount">Number of slides in document tree</param>
        /// <param name="sectionTable">Table of document sections</param>
        /// <param name="frametitleTable">Table of frame titles</param>
        /// <exception cref="PowerPointApplicationException"></exception>
        public PowerPointBuilder(string filename, string outputPath, Node document, int slideCount, List<SectionRecord> sectionTable, Dictionary<int, FrametitleRecord> frametitleTable)
        {
            if (outputPath.Length == 0)
                _filename = Path.Combine(Directory.GetCurrentDirectory(), "output", Path.GetFileNameWithoutExtension(filename));
            else
                _filename = Path.Combine(outputPath, Path.GetFileNameWithoutExtension(filename));

            _document = document;
            _slideCount = slideCount;

            _sectionTable = sectionTable ?? new List<SectionRecord>();
            _frametitleTable = frametitleTable ?? new Dictionary<int, FrametitleRecord>();

            _titlePageSettings = new Dictionary<string, List<Node>>();

            _inputFolder = Path.GetDirectoryName(filename);

            // start PowerPoint
            try
            {
                // workaround for starting PowerPoint on background
                _pptApplication = (PowerPoint.Application)Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application"));
            }
            catch
            {
                throw new PowerPointApplicationException("Couldn't start PowerPoint.");
            }

            // check office version
            if (System.Convert.ToSingle(_pptApplication.Version, System.Globalization.CultureInfo.InvariantCulture) < 12.0)
            {
                Close();    // close application
                throw new PowerPointApplicationException("You must have Office 2007 or higher!");
            }
        }

        /// <summary>
        /// Build presentation from document tree.
        /// </summary>
        public void Build()
        {
            // some ideas:
            //      - probably two tables for title settings, one local, second one global; use Dictionary<string, Node>;
            //                      clone global to local on slide start; edit global outside of slide, edit local inside of slide

            Node preambule = _document.FindFirstNode("preambule");
            if (preambule == null)
                throw new DocumentBuilderException("Couldn't build document, something went wrong. Please try again.");

            ProcessPreambule(preambule, _document.OptionalParams);

            // create new presentation without window
            _pptPresentation = _pptApplication.Presentations.Add(MsoTriState.msoFalse);

            Node body = _document.FindFirstNode("body");
            if (body == null)
                throw new DocumentBuilderException("Couldn't build document, something went wrong. Please try again.");

            ProcessBody(body);

            // todo: set type of file depending on settings window
            try
            {
                _pptPresentation.SaveAs(_filename, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
            }
            catch (Exception)
            {
                throw new DocumentBuilderException("Couldn't save output file.");
            }
            finally
            {
                // final progress change after saving
                RaiseProgress();
            }

            Messenger.Instance.SendMessage("Output saved to: \"" + _pptPresentation.FullName + "\"");
        }

        /// <summary>
        /// Raise progress counter.
        /// This method will raise progress counter based on slide count and fire (call) <see cref="Progress" /> delegate.
        /// Implemented for one pass.
        /// fixme: make this less magical :)
        /// </summary>
        public void RaiseProgress()
        {
            int step = (100 - BasicProgress) / (_slideCount);
            _currentProgress = Math.Min(_currentProgress + step, 100);

            if (Progress != null)
                Progress(_currentProgress);
        }

        /// <summary>
        /// Close PowerPoint presentations and application.
        /// </summary>
        public void Close()
        {
            if (_pptPresentation != null)
                _pptPresentation.Close();

            if (_pptApplication != null)
                _pptApplication.Quit();
        }

        /// <summary>
        /// Process document preambule part
        /// </summary>
        /// <param name="preambule">Preambule node</param>
        /// <param name="documentclassOptionals">\document class optional parameters (size)</param>
        private void ProcessPreambule(Node preambule, string documentclassOptionals)
        {
            // process preambule nodes
            foreach (Node node in preambule.Children)
            {
                switch (node.Type)
                {
                    // title page settings
                    case "author":
                    case "title":
                    case "date":
                        _titlePageSettings[node.Type] = node.Children;
                        break;
                    // unknown or invalid node -> ignore
                    default:
                        break;
                }
            }

            // process \documentclass optional parameters
            string[] parts = documentclassOptionals.Split(new char[]{','}, StringSplitOptions.RemoveEmptyEntries);
            
            foreach (string part in parts)
            {
                float size = Misc.ParseLength(part);
                if (size > 0)
                    _baseFontSize = size;
            }
        }

        /// <summary>
        /// Process document body part
        /// </summary>
        /// <param name="body">Body node</param>
        private void ProcessBody(Node body)
        {
            foreach (Node node in body.Children)
            {
                switch (node.Type)
                {
                    // title page settings
                    case "author":
                    case "title":
                    case "date":
                        _titlePageSettings[node.Type] = node.Children;
                        break;
                    // slide
                    case "slide":
                        ProcessSlide(node);
                        break;
                    // unknown or invalid node -> ignore
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Process slide content
        /// </summary>
        /// <param name="slideNode">Slide node</param>
        private void ProcessSlide(Node slideNode)
        {
            // presentation slide
            PowerPoint.Slide slide;

            // increment current slide number
            _currentSlide++;

            int passNumber = 0,
                pauseCounter = 0;

            bool titleNextPass = true,
                slideNextPass = true,
                paused = false;

            SlideBuilder slideBuilder = new SlideBuilder(_inputFolder, _currentSlide, _baseFontSize);
            TitleBuilder titleBuilder = new TitleBuilder(_baseFontSize);

            do
            {    // --- loop over all overlays

                if (paused)
                {
                    pauseCounter++;
                    paused = false;
                }

                passNumber++;
                _slideIndex++;

                // create new slide -> if slide contains title, use layout with title
                if (_frametitleTable.ContainsKey(_currentSlide))
                {
                    FrametitleRecord record = _frametitleTable[_currentSlide];

                    // check if slide has title for current pass
                    if (record.SubtitleOverlaySet.Count == 0 && record.Subtitle != null
                        || record.TitleOverlaySet.Count == 0 && record.Title != null
                        || record.SubtitleOverlaySet.Count != 0 && record.SubtitleOverlaySet.Contains(passNumber) && record.Subtitle != null
                        || record.TitleOverlaySet.Count != 0 && record.TitleOverlaySet.Contains(passNumber) && record.Title != null)
                    {
                        slide = _pptPresentation.Slides.Add(_slideIndex, PowerPoint.PpSlideLayout.ppLayoutTitleOnly);

                        titleNextPass = titleBuilder.BuildTitle(slide.Shapes[1], _frametitleTable[_currentSlide], passNumber, pauseCounter, out paused);

                        if (paused)
                            continue;

                        slideNextPass = slideBuilder.BuildSlide(slide, slideNode, new Dictionary<string, List<Node>>(_titlePageSettings), passNumber, pauseCounter, out paused);

                        continue;
                    }
                }
                
                slide = _pptPresentation.Slides.Add(_slideIndex, PowerPoint.PpSlideLayout.ppLayoutBlank);

                slideNextPass = slideBuilder.BuildSlide(slide, slideNode, new Dictionary<string, List<Node>>(_titlePageSettings), passNumber, pauseCounter, out paused);
    
            } while (!titleNextPass || !slideNextPass); // --- end loop over all overlays

            // report progress
            RaiseProgress();
        }
    }
}
