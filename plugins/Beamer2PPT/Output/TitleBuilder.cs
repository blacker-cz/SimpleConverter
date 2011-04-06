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

            // generate title
            BuildTitlePart(shape, frametitle.Title, new TextFormat(16.0f));

            if (frametitle.Subtitle != null)
            {
                shape.TextFrame2.TextRange.InsertAfter("\r");

                // generate subtitle
                BuildTitlePart(shape, frametitle.Subtitle, new TextFormat(8.0f));
            }

            return _passNumber == _maxPass;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="content"></param>
        /// <param name="format"></param>
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

                // process node here (and push __format_pop node)
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
                        // todo: check overlay here
                        format.ModifyFormat(currentNode);
                        nodes.Push(rollbackNode);
                        break;
                    case "__format_pop":
                        format.RollBackFormat();
                        break;
                    case "string":
                        format.AppendText(shape, currentNode.Content as string);
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
