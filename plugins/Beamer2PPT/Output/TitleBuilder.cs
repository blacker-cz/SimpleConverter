using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    /// <summary>
    /// Class for generating slide title content
    /// </summary>
    class TitleBuilder
    {
        /// <summary>
        /// Number of current pass on slide (used for overlay)
        /// </summary>
        private int _passNumber;

        /// <summary>
        /// Discovered number of maximum passes (from overlay specification and pause commands)
        /// </summary>
        private int _maxPass = 1;

        /// <summary>
        /// Number of processed pause commands (globally)
        /// </summary>
        private int _pauseCounter;

        /// <summary>
        /// Number of processed pause commands (locally)
        /// </summary>
        private int _localPauseCounter;

        /// <summary>
        /// Base font size
        /// </summary>
        private float _baseFontSize;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="baseFontSize">Base font size (optional)</param>
        public TitleBuilder(float baseFontSize = 11.0f)
        {
            _baseFontSize = baseFontSize;
        }

        /// <summary>
        /// Build (generate) title content
        /// </summary>
        /// <param name="shape">Title shape</param>
        /// <param name="frametitle">Frame title record</param>
        /// <param name="passNumber">Number of current pass</param>
        /// <param name="pauseCounter">Number of processed pause commands</param>
        /// <param name="paused">output - processing of title was paused</param>
        /// <returns>true if title is complete; false if needs another pass</returns>
        public bool BuildTitle(PowerPoint.Shape shape, FrametitleRecord frametitle, int passNumber, int pauseCounter, out bool paused)
        {
            _passNumber = passNumber;
            _pauseCounter = pauseCounter;
            _localPauseCounter = 0;

            // prepare shape
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignLeft;
            shape.TextFrame2.TextRange.ParagraphFormat.SpaceWithin = 0.8f;
            shape.TextFrame2.VerticalAnchor = MsoVerticalAnchor.msoAnchorTop;
            shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeShapeToFitText;
            shape.ScaleWidth(1.085f, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromMiddle);

            // setup overlays for title
            _maxPass = Math.Max(Misc.MaxOverlay(frametitle.TitleOverlaySet), _maxPass);
            _maxPass = Math.Max(Misc.MaxOverlay(frametitle.SubtitleOverlaySet), _maxPass);
            
            // generate title
            TextFormat textformat = new TextFormat(_baseFontSize);
            textformat.ModifyFormat(new Node("Large")); // in beamer title font size is same as \Large
            paused = !BuildTitlePart(shape, frametitle.Title, textformat);

            if (!paused && frametitle.Subtitle != null)
            {
                shape.TextFrame2.TextRange.InsertAfter("\r");

                // generate subtitle
                textformat = new TextFormat(_baseFontSize);
                textformat.ModifyFormat(new Node("footnotesize"));  // in beamer subtitle font size is same as \footnotesize
                paused = !BuildTitlePart(shape, frametitle.Subtitle, textformat);
            }

            return _passNumber >= _maxPass;
        }

        /// <summary>
        /// Build (generate) title part.
        /// </summary>
        /// <param name="shape">Prepared title shape</param>
        /// <param name="content">Title content</param>
        /// <param name="format">Text format</param>
        /// <returns>true if completed; false if paused</returns>
        private bool BuildTitlePart(PowerPoint.Shape shape, List<Node> content, TextFormat format)
        {
            if (content.Count == 0)   // ignore empty node
                return true;

            Stack<Node> nodes = new Stack<Node>();

            // copy content to stack
            foreach (Node item in content.Reverse<Node>())
            {
                nodes.Push(item);
            }

            Node currentNode;
            Node rollbackNode = new Node("__format_pop");

            // process nodes on stack
            while (nodes.Count != 0)
            {
                currentNode = nodes.Pop();

                // process node depending on its type
                switch (currentNode.Type)
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
                        int min = currentNode.OverlayList.Count != 0 ? currentNode.OverlayList.Min() : int.MaxValue;
                        _maxPass = Math.Max(Misc.MaxOverlay(currentNode.OverlayList), _maxPass);    // set maximal number of passes from overlay specification
                        if (currentNode.OverlayList.Count == 0 || currentNode.OverlayList.Contains(_passNumber) || min < 0 && Math.Abs(min) < _passNumber)
                        {
                            format.ModifyFormat(currentNode);
                            nodes.Push(rollbackNode);
                        }
                        break;
                    case "__format_pop":    // special node -> pop formatting from stack
                        format.RollBackFormat();
                        break;
                    case "string":
                        format.AppendText(shape, currentNode.Content as string);
                        break;
                    case "paragraph":
                        format.AppendText(shape, "\r");
                        break;
                    case "today":
                        shape.TextFrame.TextRange.InsertDateTime(PowerPoint.PpDateTimeFormat.ppDateTimeFigureOut, MsoTriState.msoTrue);
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
                    default: // unknown node -> ignore
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
    }
}
