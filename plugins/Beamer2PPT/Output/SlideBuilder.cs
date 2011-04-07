using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Text.RegularExpressions;

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
        private int _maxPass;

        /// <summary>
        /// Number of processed pause commands (globally)
        /// </summary>
        private int _pauseCounter;

        /// <summary>
        /// Number of processed pause commands (locally)
        /// </summary>
        private int _localPauseCounter;

        /// <summary>
        /// Number of processed pause commands backup start value (includes number of pauses from title)
        /// </summary>
        private int _localPauseCounterStart;

        /// <summary>
        /// Flag if BuildSlide method was called at least once
        /// </summary>
        private bool _called = false;

        /// <summary>
        /// Base font size
        /// </summary>
        private float _baseFontSize;

        /// <summary>
        /// Bottom of lowermost shape
        /// </summary>
        private float _bottomShapeBorder;

        /// <summary>
        /// Text format
        /// </summary>
        private TextFormat _format;

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
        /// <param name="pauseCounter">Number of used pauses</param>
        /// <param name="paused">output - processing of slide content was paused</param>
        /// <returns>true if slide is complete; false if needs another pass</returns>
        public bool BuildSlide(PowerPoint.Slide slide, Node slideNode, Dictionary<string, List<Node>> titlesettings, int passNumber, int pauseCounter, out bool paused)
        {
            _slide = slide;
            _titlesettings = titlesettings;
            _passNumber = passNumber;
            _pauseCounter = pauseCounter;
            _format = new TextFormat(_baseFontSize);

            if (!_called)
            {
                // because title is processed before slide and can contain pause command, we need to setup internal counters according to title passed at first processing
                _localPauseCounter = pauseCounter;
                _localPauseCounterStart = pauseCounter;
                _maxPass = passNumber;
                _called = true;
            }
            else
            {
                _localPauseCounter = _localPauseCounterStart;
            }

            // concept:
            //      iterate through nodes
            //      save font settings on stack (when entering - push new setting to stack; when leaving font settings node - pop from stack)
            //      if node is string - append to current shape
            //      if node is table/image or another shape-like object, process them separatedly
            //      at least one method for table processing and one method for image processing

            UpdateBottomShapeBorder();

            paused = !ProcessSlideContent(slideNode);

            return _passNumber >= _maxPass;
        }

        /// <summary>
        /// Process slide content
        /// </summary>
        /// <param name="slideNode">Slide content node</param>
        /// <returns>true if completed; false if paused</returns>
        private bool ProcessSlideContent(Node slideNode)
        {
            // note: width of slide content area is 648.0
            if (slideNode.Children.Count == 0)   // ignore empty node
                return true;

            Stack<Node> nodes = new Stack<Node>();

            // copy content to stack
            foreach (Node item in slideNode.Children.Reverse<Node>())
            {
                nodes.Push(item);
            }

            Node currentNode;
            Node rollbackNode = new Node("__format_pop");

            PowerPoint.Shape shape = null;

            // process nodes on stack
            while (nodes.Count != 0)
            {
                currentNode = nodes.Pop();

                // process node depending on its type
                switch (currentNode.Type)
                {
                    case "string":
                        if (shape == null)
                        {
                            UpdateBottomShapeBorder(true);
                            shape = _slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 36.0f, _bottomShapeBorder + 5.0f, 648.0f, 10.0f);
                        }
                        _format.AppendText(shape, currentNode.Content as string);
                        break;
                    case "paragraph":
                        if(shape != null)
                            _format.AppendText(shape, "\r");
                        break;
                    case "pause":
                        _localPauseCounter++;

                        if (_localPauseCounter > _pauseCounter)
                        {
                            if (_passNumber == _maxPass)    // increase number of passes
                                _maxPass++;
                            return false;
                        }
                        break;
                    case "today":
                        // todo: check shape existence
                        shape.TextFrame.TextRange.InsertDateTime(PowerPoint.PpDateTimeFormat.ppDateTimeFigureOut, MsoTriState.msoTrue);
                        break;
                    case "image":
                        UpdateBottomShapeBorder(true);
                        shape = null;
                        // todo: process in separate method
                        // reposition shapes and next text shape create on right side of image, then everything after first line copy to new shape below image
                        break;
                    case "bulletlist":
                    case "numberedlist":
                        UpdateBottomShapeBorder(true);
                        shape = null;
                        // todo: process in separate method or here?
                        break;
                    case "table":
                        break;
                    case "descriptionlist":
                        // todo: implement this probably as simple table
                        break;
                    default: // other -> check for simple formats
                        SimpleTextFormat(nodes, currentNode);
                        break;
                }

                if (currentNode.Children == null)
                    continue;

                // push child nodes to stack
                foreach (Node item in currentNode.Children.Reverse<Node>())
                {
                    nodes.Push(item);
                }
            }

            return true;
        }

        /// <summary>
        /// Compute and save bottom position of the lowermost shape
        /// </summary>
        /// <param name="autoTrim">Auto trim content in TextFrame (if available)</param>
        private void UpdateBottomShapeBorder(bool autoTrim = false)
        {
            if (_slide.Shapes.Count != 0)
            {
                foreach (PowerPoint.Shape shape in _slide.Shapes)
                {
                    if (autoTrim && shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame2.HasText == MsoTriState.msoTrue)
                    {
                        // TextRange.TrimText() method is useless because it doesn't actually remove whitespaces from text range but returns its copy
                        // first compute number of whitespace characters at the end of shape
                        int size = shape.TextFrame2.TextRange.Text.Length - shape.TextFrame2.TextRange.Text.TrimEnd().Length;
                        if (size > 0)   // then if there is more then zero of these characters -> delete them
                            shape.TextFrame2.TextRange.Characters[1 + shape.TextFrame2.TextRange.Text.Length - size, size].Delete();
                    }

                    // compute bottom of the lowermost shape
                    _bottomShapeBorder = Math.Max(_bottomShapeBorder, (shape.Height + shape.Top));
                }
            }
        }

        /// <summary>
        /// Check if TextRange object ends with new line (and ignore other whitespace characters during check)
        /// </summary>
        /// <param name="textRange">Text range</param>
        /// <returns>true if yes; false if otherwise</returns>
        private bool EndsWithNewLine(TextRange2 textRange)
        {
            Regex reg = new Regex("\r[\t ]*$");
            return reg.IsMatch(textRange.Text);
        }

        /// <summary>
        /// Process simple text formatting
        /// </summary>
        /// <param name="nodes">Nodes stack - used for pushing format rollback node</param>
        /// <param name="node">Current node</param>
        private void SimpleTextFormat(Stack<Node> nodes, Node node)
        {
            Node rollbackNode = new Node("__format_pop");

            switch (node.Type)
            {
                case "bold":
                case "italic":
                case "underline":
                case "smallcaps":
                case "typewriter":
                case "color":
                case "tiny":
                case "scriptsize":
                case "footnotesize":
                case "small":
                case "normalsize":
                case "large":
                case "Large":
                case "LARGE":
                case "huge":
                case "Huge":
                    // check overlay settings
                    int min = node.OverlayList.Count != 0 ? node.OverlayList.Min() : int.MaxValue;
                    _maxPass = Math.Max(Misc.MaxOverlay(node.OverlayList), _maxPass);    // set maximal number of passes from overlay specification
                    if (node.OverlayList.Count == 0 || node.OverlayList.Contains(_passNumber) || min < 0 && Math.Abs(min) < _passNumber)
                    {
                        _format.ModifyFormat(node);
                        nodes.Push(rollbackNode);
                    }
                    break;
                case "__format_pop":    // special node -> pop formatting from stack
                    _format.RollBackFormat();
                    break;
                default:    // unknown node -> ignore
                    break;
            }
        }
    }
}
