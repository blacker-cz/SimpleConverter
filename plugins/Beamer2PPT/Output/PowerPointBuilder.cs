﻿using System;
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
        private int _currentProgress;

        /// <summary>
        /// Filename of output file (without extension)
        /// </summary>
        private string _filename;

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
        /// Preambule settings
        /// </summary>
        private PreambuleSettings _preambuleSettings;

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
        /// <exception cref="PowerPointApplicationException"></exception>
        public PowerPointBuilder()
        {
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
        /// <param name="filename">Filename of input file (used for output file)</param>
        /// <param name="outputPath">Output directory</param>
        /// <param name="document">Document tree</param>
        /// <param name="slideCount">Number of slides in document tree</param>
        /// <param name="sectionTable">Table of document sections</param>
        /// <param name="frametitleTable">Table of frame titles</param>
        public void Build(string filename, string outputPath, Node document, int slideCount, List<SectionRecord> sectionTable, Dictionary<int, FrametitleRecord> frametitleTable)
        {

            #region Initialize internal variables
            if (outputPath.Length == 0)
            {
                _filename = Path.Combine(Directory.GetCurrentDirectory(), "output", Path.GetFileNameWithoutExtension(filename));

                try
                {
                    // if output directory doesn't exist then create it
                    if (!Directory.Exists(Path.Combine(Directory.GetCurrentDirectory(), "output")))
                        Directory.CreateDirectory(Path.Combine(Directory.GetCurrentDirectory(), "output"));
                }
                catch (Exception)
                {
                    throw new PowerPointApplicationException("Couldn't create default output directory.");
                }
            }
            else
            {
                _filename = Path.Combine(outputPath, Path.GetFileNameWithoutExtension(filename));

                try
                {
                    // if output directory doesn't exist then create it
                    if (!Directory.Exists(outputPath))
                        Directory.CreateDirectory(outputPath);
                }
                catch (Exception)
                {
                    throw new PowerPointApplicationException("Couldn't create output directory.");
                }
            }

            _document = document;
            _slideCount = slideCount;

            _sectionTable = sectionTable ?? new List<SectionRecord>();
            _frametitleTable = frametitleTable ?? new Dictionary<int, FrametitleRecord>();

            _preambuleSettings = new PreambuleSettings(Path.GetDirectoryName(filename));

            _currentSlide = 0;
            _slideIndex = 0;

            _currentProgress = BasicProgress;

            #endregion // Initialize internal variables

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

            try
            {
                _pptPresentation.SaveAs(_filename, Settings.Instance.SaveAs);
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

            // print save message
            switch (Settings.Instance.SaveAs)
            {
                case Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault:
                case Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation:
                case Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsPresentation:
                    Messenger.Instance.SendMessage(System.IO.Path.GetFileName(filename) + " - Output saved to: \"" + _pptPresentation.FullName + "\"");
                    break;
                case Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsPDF:
                    Messenger.Instance.SendMessage(System.IO.Path.GetFileName(filename) + " - Output saved to: \"" + _filename + ".pdf\"");
                    break;
                default:
                    Messenger.Instance.SendMessage(System.IO.Path.GetFileName(filename) + " - Output saved to output directory.");
                    break;
            }

            try
            {
                _pptPresentation.Close();
                _pptPresentation = null;
            }
            catch { }
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
            {
                _pptPresentation.Close();
                _pptPresentation = null;
            }

            if (_pptApplication != null)
            {
                _pptApplication.Quit();
                _pptApplication = null;
            }
        }

        /// <summary>
        /// Process document preambule part
        /// </summary>
        /// <param name="preambule">Preambule node</param>
        /// <param name="documentclassOptionals">\document class optional parameters (size)</param>
        private void ProcessPreambule(Node preambule, string documentclassOptionals)
        {
            // parse preambule
            _preambuleSettings.Parse(preambule);

            // process \documentclass optional parameters
            string[] parts = documentclassOptionals.Split(new char[]{','}, StringSplitOptions.RemoveEmptyEntries);
            
            foreach (string part in parts)
            {
                float size = Misc.ParseLength(part);
                if (size > 0)
                    _baseFontSize = Settings.Instance.AdjustSize ? size / 2 : size;
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
                    case "institute":
                    case "date":
                        _preambuleSettings.TitlepageSettings[node.Type] = node.Children;
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

            int maxPass = 0;

            bool titleNextPass = true,
                slideNextPass = true,
                paused = false;

            SlideBuilder slideBuilder = new SlideBuilder(_preambuleSettings, _currentSlide, _baseFontSize);
            TitleBuilder titleBuilder = new TitleBuilder(_baseFontSize);

            // list of sub-slides
            List<PowerPoint.Slide> subSlides = new List<PowerPoint.Slide>();

            // slide settings
            SlideSettings slideSettings = new SlideSettings(slideNode.OptionalParams);

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
                        || Misc.ShowOverlay(passNumber, record.SubtitleOverlaySet, ref maxPass) && record.Subtitle != null
                        || Misc.ShowOverlay(passNumber, record.TitleOverlaySet, ref maxPass) && record.Title != null)
                    {
                        slide = _pptPresentation.Slides.Add(_slideIndex, PowerPoint.PpSlideLayout.ppLayoutTitleOnly);
                        subSlides.Add(slide);   // add slide to list of sub-slides

                        titleNextPass = titleBuilder.BuildTitle(slide.Shapes.Title, _frametitleTable[_currentSlide], passNumber, pauseCounter, out paused);

                        if (paused)
                            continue;

                        slideNextPass = slideBuilder.BuildSlide(slide, slideNode, passNumber, pauseCounter, out paused);

                        continue;
                    }
                }
                
                slide = _pptPresentation.Slides.Add(_slideIndex, PowerPoint.PpSlideLayout.ppLayoutBlank);
                subSlides.Add(slide);   // add slide to list of sub-slides

                slideNextPass = slideBuilder.BuildSlide(slide, slideNode, passNumber, pauseCounter, out paused);
    
            } while (!titleNextPass || !slideNextPass); // --- end loop over all overlays

            // change slide content vertical align
            if (slideSettings.ContentAlign != SlideSettings.Align.TOP)
            {
                float bottom = float.MinValue;

                // compute maximal top and bottom of slide content
                foreach (PowerPoint.Slide subSlide in subSlides)
                {
                    bottom = Math.Max(bottom, subSlide.Shapes[subSlide.Shapes.Count].Top + subSlide.Shapes[subSlide.Shapes.Count].Height);
                }

                float change;
                
                if (slideSettings.ContentAlign == SlideSettings.Align.BOTTOM)
                    change = (540.0f - bottom) - 10.0f;
                else
                    change = (540.0f - bottom) / 2.5f;

                foreach (PowerPoint.Slide subSlide in subSlides)
                {
                    foreach (PowerPoint.Shape shape in subSlide.Shapes)
                    {
                        if (subSlide.Shapes.HasTitle != MsoTriState.msoTrue || shape != subSlide.Shapes.Title)
                        {
                            shape.Top = shape.Top + change;
                        }
                    }
                }
            }

            // report progress
            RaiseProgress();
        }
    }
}
