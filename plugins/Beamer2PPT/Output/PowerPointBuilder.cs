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
        /// Magic number for progress after document parsings
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

        #endregion // Private variables

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="filename">Filename of input file (used for output file)</param>
        /// <param name="outputPath">Output directory</param>
        /// <param name="document">Document tree</param>
        /// <param name="slideCount">Number of slides in document tree</param>
        public PowerPointBuilder(string filename, string outputPath, Node document, int slideCount)
        {
            if (outputPath.Length == 0)
                _filename = Path.Combine(Directory.GetCurrentDirectory(), "output", Path.GetFileNameWithoutExtension(filename));
            else
                _filename = Path.Combine(outputPath, Path.GetFileNameWithoutExtension(filename));

            _document = document;
            _slideCount = slideCount;

            // start PowerPoint
            try
            {
                // workaround for starting PowerPoint on background
                _pptApplication = (PowerPoint.Application) Activator.CreateInstance(Type.GetTypeFromProgID("PowerPoint.Application"));
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
        /// <returns>true if no errors occured, false otherwise</returns>
        public bool Build()
        {
            // some ideas:
            // - two runs through document tree
            //      - in first run setup content table, frametitles and subtitles table, maybe references (if implemented), packages, theme etc.
            //      - in second run build output document
            throw new NotImplementedException();
        }

        /// <summary>
        /// Raise progress counter.
        /// This method will raise progress counter based on slide count and fire (call) <see cref="Progress" /> delegate.
        /// </summary>
        public void RaiseProgress()
        {
            int step = (100 - BasicProgress) / _slideCount;
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
    }
}
