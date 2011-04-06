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
        /// Build (generate) title content
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="frametitle"></param>
        /// <param name="passNumber"></param>
        /// <returns>true if title is complete; false if needs another pass</returns>
        public bool BuildTitle(PowerPoint.Shape shape, FrametitleRecord frametitle, int passNumber)
        {
            _passNumber = passNumber;

            // prepare shape
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignLeft;
            shape.TextFrame2.TextRange.ParagraphFormat.SpaceWithin = 0.75f;
            shape.TextFrame2.VerticalAnchor = MsoVerticalAnchor.msoAnchorTop;
            shape.TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeShapeToFitText;
            shape.ScaleWidth(1.085f, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromMiddle);

            // setup overlays for title
            _maxPass = Math.Max(Misc.MaxOverlay(frametitle.TitleOverlaySet), _maxPass);
            _maxPass = Math.Max(Misc.MaxOverlay(frametitle.SubtitleOverlaySet), _maxPass);
            
            // generate title
            BuildTitlePart(shape, frametitle.Title, new TextFormat(16.0f));

            if (frametitle.Subtitle != null)
            {
                shape.TextFrame2.TextRange.InsertAfter("\r");

                // generate subtitle
                BuildTitlePart(shape, frametitle.Subtitle, new TextFormat(8.0f));
            }

            return _passNumber >= _maxPass;
        }

        /// <summary>
        /// Build (generate) title part.
        /// </summary>
        /// <param name="shape">Prepared title shape</param>
        /// <param name="content">Title content</param>
        /// <param name="format">Text format</param>
        private void BuildTitlePart(PowerPoint.Shape shape, List<Node> content, TextFormat format)
        {
            if (content.Count == 0)   // ignore empty node
                return;

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
                        // todo: test this
                        int min = currentNode.OverlayList.Count != 0 ? currentNode.OverlayList.Min() : int.MaxValue;
                        if (currentNode.OverlayList.Count == 0 || currentNode.OverlayList.Contains(_passNumber) || min < 0 && Math.Abs(min) < _passNumber)
                        {
                            _maxPass = Math.Max(Misc.MaxOverlay(currentNode.OverlayList), _maxPass);
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
                        // todo: keep number of processed pauses (local) if passNumber < pauseCount -> throw pause exception (catch in PowerPointBuilder, and continue - before processing content)
                        // else ignore
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
        }
    }
}
