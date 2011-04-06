using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    class SlideBuilder
    {
        /// <summary>
        /// Currently created slide
        /// </summary>
        private PowerPoint.Slide _slide;

        /// <summary>
        /// Table with title settings (used for \maketitle)
        /// </summary>
        private Dictionary<string, List<Node>> _titlesettings;

        /// <summary>
        /// Number of current slide
        /// </summary>
        private int _slideNumber;

        /// <summary>
        /// Number of current pass on slide (used for overlay)
        /// </summary>
        private int _passNumber;

        /// <summary>
        /// Discovered number of maximum passes (from overlay specification and pause commands)
        /// todo: with overlay set to max; with pause set to current pass + 1, only if > _maxPass
        /// </summary>
        private int _maxPass = 1;

        /// <summary>
        /// Number of processed pause commands
        /// </summary>
        private int _pausedCounter;

        /// <summary>
        /// Base font size
        /// </summary>
        private float _baseFontSize;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="slideNumber">Number of currently generated slide</param>
        /// <param name="baseFontSize">Base font size (optional)</param>
        public SlideBuilder(int slideNumber, float baseFontSize = 10.0f)
        {
            _slideNumber = slideNumber;
            _baseFontSize = baseFontSize;
        }

        /// <summary>
        /// Build slide content
        /// </summary>
        /// <param name="slide">Slide in PowerPoint presentation</param>
        /// <param name="slideNode">Node containing content of slide</param>
        /// <param name="titlesettings">Title settings table (used for \maketitle command)</param>
        /// <param name="passNumber">Number of current pass (used for overlays)</param>
        /// <returns>true if slide is complete; false if needs another pass</returns>
        public bool BuildSlide(PowerPoint.Slide slide, Node slideNode, Dictionary<string, List<Node>> titlesettings, int passNumber, int pausedCounter, out bool paused)
        {
            _slide = slide;
            _titlesettings = titlesettings;
            _passNumber = passNumber;
            _pausedCounter = pausedCounter;

            // concept:
            //      iterate through nodes
            //      save font settings on stack (when entering - push new setting to stack; when leaving font settings node - pop from stack)
            //      if node is string - append to current shape
            //      if node is table/image or another shape-like object, process them separatedly
            //      at least one method for table processing and one method for image processing

            paused = false;

            return _passNumber >= _maxPass;
        }
    }
}
