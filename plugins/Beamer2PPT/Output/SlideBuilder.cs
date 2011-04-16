using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Text.RegularExpressions;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    /// <summary>
    /// Class for generating slide content
    /// </summary>
    class SlideBuilder
    {
        #region Private variables

        /// <summary>
        /// Currently created slide
        /// </summary>
        private PowerPoint.Slide _slide;

        /// <summary>
        /// Table with title settings (used for \titlepage)
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
        /// Preambule settings
        /// </summary>
        private PreambuleSettings _preambuleSettings;

        /// <summary>
        /// Queue for nested elements generated at bottom of slide
        /// </summary>
        private Queue<Node> _postProcessing;

        #endregion // Private variables

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="preambuleSettings">Preambule settings</param>
        /// <param name="slideNumber">Number of currently generated slide</param>
        /// <param name="baseFontSize">Base font size (optional)</param>
        public SlideBuilder(PreambuleSettings preambuleSettings, int slideNumber, float baseFontSize = 11.0f)
        {
            _slideNumber = slideNumber;
            _baseFontSize = baseFontSize;
            _preambuleSettings = preambuleSettings;
        }

        /// <summary>
        /// Build slide content
        /// </summary>
        /// <param name="slide">Slide in PowerPoint presentation</param>
        /// <param name="slideNode">Node containing content of slide</param>
        /// <param name="passNumber">Number of current pass (used for overlays)</param>
        /// <param name="pauseCounter">Number of used pauses</param>
        /// <param name="paused">output - processing of slide content was paused</param>
        /// <returns>true if slide is complete; false if needs another pass</returns>
        public bool BuildSlide(PowerPoint.Slide slide, Node slideNode, int passNumber, int pauseCounter, out bool paused)
        {
            _slide = slide;
            _passNumber = passNumber;
            _pauseCounter = pauseCounter;
            // copy title settings
            _titlesettings = new Dictionary<string,List<Node>>(_preambuleSettings.TitlepageSettings);
            _format = new TextFormat(_baseFontSize);
            _postProcessing = new Queue<Node>();

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

            UpdateBottomShapeBorder();

            paused = !ProcessSlideContent(slideNode);

            if (!Settings.Instance.NestedAsText)    // extract nested elements
            {
                _localPauseCounter = int.MinValue;  // ignore pauses in post processed shapes
                PostProcessing();
            }

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

            PowerPoint.Shape shape = null;

            // skip expanding of child nodes to stack
            bool skip;

            // list of shapes used for reshaper
            List<PowerPoint.Shape> reshapeShapes = new List<PowerPoint.Shape>();

            // process nodes on stack
            while (nodes.Count != 0)
            {
                currentNode = nodes.Pop();
                skip = false;

                if(reshapeShapes.Count > 1)
                    Reshaper(reshapeShapes);

                // process node depending on its type
                switch (currentNode.Type)
                {
                    case "string":
                        if (shape == null)
                        {
                            UpdateBottomShapeBorder(true);
                            _format.Invalidate();
                            shape = _slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 36.0f, _bottomShapeBorder + 5.0f, 648.0f, 10.0f);
                            string tmp = currentNode.Content as string;
                            if (Misc.StartsWithNewLine(tmp))    // if new shape starts with new line, then call reshaper on previous shapes
                            {
                                Reshaper(reshapeShapes);
                                reshapeShapes.Clear();
                                tmp = tmp.TrimStart();
                                if (tmp.Length == 0)    // only empty paragraph -> don't start new shape
                                {
                                    shape.Delete();
                                    shape = null;
                                }
                                else
                                {
                                    _format.AppendText(shape, tmp);
                                }
                            }
                            else
                            {
                                _format.AppendText(shape, tmp);
                            }
                            reshapeShapes.Add(shape);
                        }
                        else
                        {
                            _format.AppendText(shape, currentNode.Content as string);
                        }
                        break;
                    case "paragraph":
                        if(shape != null && shape.HasTextFrame == MsoTriState.msoTrue && !Misc.EndsWithNewLine(shape.TextFrame2.TextRange.Text))
                            _format.AppendText(shape, "\r");

                        // also do reshaping
                        Reshaper(reshapeShapes);
                        reshapeShapes.Clear();

                        break;
                    case "pause":
                        if (Pause())
                            return false;
                        break;
                    case "today":
                        if (shape == null)
                        {
                            UpdateBottomShapeBorder(true);
                            _format.Invalidate();
                            shape = _slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 36.0f, _bottomShapeBorder + 5.0f, 648.0f, 10.0f);
                            reshapeShapes.Add(shape);
                        }
                        shape.TextFrame.TextRange.InsertDateTime(PowerPoint.PpDateTimeFormat.ppDateTimeFigureOut, MsoTriState.msoTrue);
                        break;
                    case "image":
                        // previous shape ends with new paragraph - image will start on new line
                        if (shape != null && shape.HasTextFrame == MsoTriState.msoTrue && Misc.EndsWithNewLine(shape.TextFrame2.TextRange.Text))
                        {
                            Reshaper(reshapeShapes);
                            reshapeShapes.Clear();
                        }
                        UpdateBottomShapeBorder(true);
                        PowerPoint.Shape imageShape;

                        GenerateImage(currentNode, out imageShape);
                        reshapeShapes.Add(imageShape);

                        shape = null;
                        break;
                    case "bulletlist":
                    case "numberedlist":
                        skip = true;

                        Reshaper(reshapeShapes);
                        reshapeShapes.Clear();
                        
                        UpdateBottomShapeBorder(true);
                        shape = _slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 36.0f, _bottomShapeBorder + 5.0f, 648.0f, 10.0f);

                        if (!GenerateList(currentNode.Children, shape, 1, currentNode.Type == "bulletlist" ? MsoBulletType.msoBulletUnnumbered : MsoBulletType.msoBulletNumbered))
                            return false;

                        shape = null;
                        break;
                    case "table":
                        skip = true;
                        // previous shape ends with new paragraph - table will start on new line
                        if (shape != null && shape.HasTextFrame == MsoTriState.msoTrue && Misc.EndsWithNewLine(shape.TextFrame2.TextRange.Text))
                        {
                            Reshaper(reshapeShapes);
                            reshapeShapes.Clear();
                        }
                        UpdateBottomShapeBorder(true);
                        PowerPoint.Shape tableShape;

                        if (!GenerateTable(currentNode, out tableShape))
                        {
                            reshapeShapes.Add(tableShape);  // reshape before ending this pass
                            Reshaper(reshapeShapes);

                            return false;   // table processing was paused
                        }

                        if (tableShape == null) // table wasn't generated so call reshaper and next shape will start at new "line"
                        {
                            Reshaper(reshapeShapes);
                            reshapeShapes.Clear();
                        }
                        else
                        {
                            reshapeShapes.Add(tableShape);
                        }

                        shape = null;
                        break;
                    case "descriptionlist":
                        skip = true;

                        Reshaper(reshapeShapes);
                        reshapeShapes.Clear();
                        
                        UpdateBottomShapeBorder(true);

                        if (!GenerateDescriptionList(currentNode.Children))
                            return false;

                        shape = null;
                        break;
                    case "titlepage":

                        Reshaper(reshapeShapes);
                        reshapeShapes.Clear();
                        
                        UpdateBottomShapeBorder(true);

                        GenerateTitlePage();

                        shape = null;
                        break;
                    default: // other -> check for simple formats
                        SimpleTextFormat(nodes, currentNode);
                        break;
                }

                if (currentNode.Children == null || skip)
                    continue;

                // push child nodes to stack
                foreach (Node item in currentNode.Children.Reverse<Node>())
                {
                    nodes.Push(item);
                }
            }

            // final reshaper call on slide
            Reshaper(reshapeShapes);

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
                    if (autoTrim)
                    {
                        Misc.TrimShape(shape);
                    }

                    // compute bottom of the lowermost shape
                    _bottomShapeBorder = Math.Max(_bottomShapeBorder, (shape.Height + shape.Top));
                }
            }
            else    // no shape -> set 15pt from top of slide
            {
                _bottomShapeBorder = 15.0f;
            }
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

        /// <summary>
        /// Check if should pause on current pause command
        /// </summary>
        /// <returns>true if should; false otherwise</returns>
        private bool Pause()
        {
            _localPauseCounter++;

            if (_localPauseCounter > _pauseCounter)
            {
                if (_passNumber == _maxPass)    // increase number of passes
                    _maxPass++;
                return true;
            }

            return false;
        }

        /// <summary>
        /// Generate table from its node
        /// </summary>
        /// <param name="tableNode">Table node</param>
        /// <param name="tableShape">output - Shape of generated table (used for reshaper), null if no table was generated</param>
        /// <returns>true if completed; false if paused</returns>
        private bool GenerateTable(Node tableNode, out PowerPoint.Shape tableShape)
        {
            int rows = 0, cols = 0;
            tableShape = null;

            TabularSettings settings = TabularSettings.Parse(tableNode.Content as string);

            cols = settings.Columns.Count;

            // count table rows
            foreach (Node node in tableNode.Children)
            {
                if (node.Type == "tablerow")
                    rows++;
            } // counted number of rows can be exactly one row greater than actual value (last row is empty)

            if (cols == 0 || rows == 0) // no columns or rows -> don't create table
                return true;

            // create table shape with "rows - 1" rows but at least one row; also create table with extreme width so we can resize it down
            tableShape = _slide.Shapes.AddTable(((rows - 1) > 0 ? rows - 1 : rows), cols, 36.0f, _bottomShapeBorder + 5.0f, cols * 1000.0f);
            // style without background and borders
            tableShape.Table.ApplyStyle("2D5ABB26-0587-4C30-8999-92F81FD0307C");

            int rowCounter = 0, columnCounter = 0;

            Stack<Node> nodes = new Stack<Node>();

            Node currentNode;

            PowerPoint.Shape shape; // cell shape

            // skip expanding children to stack
            bool skip = false;

            // pause processing variables
            bool paused = false;
            int pausedAfter = 0;

            foreach (Node node in tableNode.Children)
            {
                columnCounter = 0;

                if (node.Type == "tablerow")
                {
                    rowCounter++;

                    // check if we will generate last row
                    if (rowCounter == rows && rowCounter != 1)
                    {
                        if (node.Children.Count == 1 && node.Children[0].Children.Count == 0)
                            continue;
                        else
                            tableShape.Table.Rows.Add();
                    }

                    foreach (Node rowcontent in node.Children)
                    {
                        if (rowcontent.Type == "tablecolumn" || rowcontent.Type == "tablecolumn_merged")
                        {
                            columnCounter++;

                            if (columnCounter > cols)
                                throw new DocumentBuilderException("Invalid table definition.");

                            // get cell shape
                            shape = tableShape.Table.Cell(rowCounter, columnCounter).Shape;

                            // set cell alignment
                            switch (settings.Columns[columnCounter-1].alignment)
                            {
                                case 'l':
                                    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignLeft;
                                    break;
                                case 'c':
                                    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter;
                                    break;
                                case 'r':
                                    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignRight;
                                    break;
                                case 'p':
                                    shape.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignJustify;
                                    break;
                                default:
                                    break;
                            }

                            _format.Invalidate();

                            // copy column content to stack
                            foreach (Node item in rowcontent.Children.Reverse<Node>())
                            {
                                nodes.Push(item);
                            }

                            // process nodes on stack
                            while (nodes.Count != 0)
                            {
                                currentNode = nodes.Pop();

                                skip = false;

                                // process node depending on its type
                                switch (currentNode.Type)
                                {
                                    case "string":
                                        _format.AppendText(shape, currentNode.Content as string);
                                        break;
                                    case "paragraph":
                                        _format.AppendText(shape, "\r");
                                        break;
                                    case "pause":
                                        if (!paused && Pause())
                                        {
                                            paused = true;
                                            if (columnCounter == 1 && shape.TextFrame2.TextRange.Text.Length == 0)
                                                pausedAfter = rowCounter - 1;
                                            else
                                                pausedAfter = rowCounter;
                                        }
                                        break;
                                    case "numberedlist":
                                    case "bulletlist":
                                    case "descriptionlist":
                                    case "image":
                                    case "table":
                                        if (!Settings.Instance.NestedAsText)
                                        {
                                            skip = true;
                                            _postProcessing.Enqueue(currentNode);
                                        }
                                        break;
                                    case "today":
                                        shape.TextFrame.TextRange.InsertDateTime(PowerPoint.PpDateTimeFormat.ppDateTimeFigureOut, MsoTriState.msoTrue);
                                        break;
                                    default: // other -> check for simple formats
                                        SimpleTextFormat(nodes, currentNode);
                                        break;
                                }

                                if (currentNode.Children == null || skip)
                                    continue;

                                // push child nodes to stack
                                foreach (Node item in currentNode.Children.Reverse<Node>())
                                {
                                    nodes.Push(item);
                                }
                            }

                            if (rowcontent.Type == "tablecolumn")
                            {
                                if (columnCounter == 1 && settings.Borders.Contains(0)) // first column check also for border with index 0 (left border)
                                {
                                    tableShape.Table.Rows[rowCounter].Cells[columnCounter].Borders[PowerPoint.PpBorderType.ppBorderLeft].ForeColor.RGB = 0x0;
                                    tableShape.Table.Rows[rowCounter].Cells[columnCounter].Borders[PowerPoint.PpBorderType.ppBorderLeft].DashStyle = MsoLineDashStyle.msoLineSolid;
                                }

                                if (settings.Borders.Contains(columnCounter))   // for every column set right border
                                {
                                    tableShape.Table.Rows[rowCounter].Cells[columnCounter].Borders[PowerPoint.PpBorderType.ppBorderRight].ForeColor.RGB = 0x0;
                                    tableShape.Table.Rows[rowCounter].Cells[columnCounter].Borders[PowerPoint.PpBorderType.ppBorderRight].DashStyle = MsoLineDashStyle.msoLineSolid;
                                }
                            }

                            // merge cells
                            if (rowcontent.Type == "tablecolumn_merged")
                            {
                                // merge cells here and increment columnCounter depending on number of merged cells
                                string tmp = rowcontent.Content as string;

                                int merge_count;

                                if (int.TryParse(tmp.Trim(), out merge_count))
                                {
                                    // merge cells
                                    tableShape.Table.Cell(rowCounter, columnCounter).Merge(tableShape.Table.Cell(rowCounter, columnCounter + merge_count - 1));

                                    TabularSettings mset = TabularSettings.Parse(rowcontent.OptionalParams, true);

                                    // left border
                                    if (mset.Borders.Contains(0))
                                    {
                                        tableShape.Table.Cell(rowCounter, columnCounter).Borders[PowerPoint.PpBorderType.ppBorderLeft].ForeColor.RGB = 0x0;
                                        tableShape.Table.Cell(rowCounter, columnCounter).Borders[PowerPoint.PpBorderType.ppBorderLeft].DashStyle = MsoLineDashStyle.msoLineSolid;
                                    }

                                    // right border
                                    if (mset.Borders.Contains(1))
                                    {
                                        tableShape.Table.Cell(rowCounter, columnCounter + merge_count - 1).Borders[PowerPoint.PpBorderType.ppBorderRight].ForeColor.RGB = 0x0;
                                        tableShape.Table.Cell(rowCounter, columnCounter + merge_count - 1).Borders[PowerPoint.PpBorderType.ppBorderRight].DashStyle = MsoLineDashStyle.msoLineSolid;
                                    }

                                    // set cell alignment
                                    switch (mset.Columns[0].alignment)
                                    {
                                        case 'l':
                                            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignLeft;
                                            break;
                                        case 'c':
                                            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter;
                                            break;
                                        case 'r':
                                            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignRight;
                                            break;
                                        case 'p':
                                            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignJustify;
                                            break;
                                        default:
                                            break;
                                    }

                                    // skip merged columns
                                    columnCounter += merge_count - 1;
                                }
                            }
                        }
                    }
                }
                else if(node.Type == "hline")
                {
                    if (rowCounter == 0)
                    {
                        tableShape.Table.Rows[1].Cells.Borders[PowerPoint.PpBorderType.ppBorderTop].ForeColor.RGB = 0x0;
                        tableShape.Table.Rows[1].Cells.Borders[PowerPoint.PpBorderType.ppBorderTop].DashStyle = MsoLineDashStyle.msoLineSolid;
                    }
                    else
                    {
                        tableShape.Table.Rows[rowCounter].Cells.Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = 0x0;
                        tableShape.Table.Rows[rowCounter].Cells.Borders[PowerPoint.PpBorderType.ppBorderBottom].DashStyle = MsoLineDashStyle.msoLineSolid;
                    }
                }
                else if (node.Type == "cline")
                {
                    Regex regex = new Regex(@"^([0-9]+)-([0-9]+)$", RegexOptions.IgnoreCase);

                    string range = node.Content as string;

                    Match match = regex.Match(range.Trim());

                    if (match.Success)
                    {
                        int x, y;

                        if (int.TryParse(match.Groups[1].Value, out x) && int.TryParse(match.Groups[2].Value, out y))
                        {
                            for (int i = Math.Min(x,y); i <= Math.Max(x,y); i++)
                            {
                                if (rowCounter == 0)
                                {
                                    tableShape.Table.Rows[1].Cells[i].Borders[PowerPoint.PpBorderType.ppBorderTop].ForeColor.RGB = 0x0;
                                    tableShape.Table.Rows[1].Cells[i].Borders[PowerPoint.PpBorderType.ppBorderTop].DashStyle = MsoLineDashStyle.msoLineSolid;
                                }
                                else
                                {
                                    tableShape.Table.Rows[rowCounter].Cells[i].Borders[PowerPoint.PpBorderType.ppBorderBottom].ForeColor.RGB = 0x0;
                                    tableShape.Table.Rows[rowCounter].Cells[i].Borders[PowerPoint.PpBorderType.ppBorderBottom].DashStyle = MsoLineDashStyle.msoLineSolid;
                                }
                            }
                        }
                    }
                }
            }

            // resize table
            Misc.AutoFitColumn(tableShape, settings);

            // if processing was paused remove all lines after pause commands (columns are not supported yet)
            if (paused)
            {
                if (pausedAfter == 0)
                {
                    tableShape.Delete();
                    tableShape = null;
                }
                else
                {
                    for (int i = tableShape.Table.Rows.Count; i > pausedAfter; i--)
                    {
                        tableShape.Table.Rows[i].Delete();
                    }
                }

                return false;
            }

            return true;
        }

        /// <summary>
        /// Insert image from its node
        /// </summary>
        /// <param name="imageNode">Image node</param>
        /// <param name="imageShape">output - Shape of inserted image (used for reshaper), null if no image was inserted</param>
        private void GenerateImage(Node imageNode, out PowerPoint.Shape imageShape)
        {
            imageShape = null;

            // find inserted image
            string imagePath = Misc.FindImage(imageNode.Content as string, _preambuleSettings.GraphicsPath);
            if (imagePath == null)
            {
                throw new DocumentBuilderException("Image '" + imageNode.Content + "' not found!");
            }

            try     // just to be sure
            {
                imageShape = _slide.Shapes.AddPicture(imagePath, MsoTriState.msoFalse, MsoTriState.msoTrue, 36.0f, _bottomShapeBorder + 5.0f);
            }
            catch (Exception)
            {
                // throw image not found exception
                throw new DocumentBuilderException("Image '" + imageNode.Content + "' not found!");
            }

            float width = 0, height = 0, scale = 0;

            // parse optional parameters
            string[] optParams = imageNode.OptionalParams.Replace(" ", "").Split(new Char[] {','}, StringSplitOptions.RemoveEmptyEntries);
            foreach (string item in optParams)
            {
                string[] pair = item.Split('=');

                switch (pair[0])
                {
                    case "width":
                        if (pair.Length > 1)
                            width = Misc.ParseLength(pair[1]);
                        break;
                    case "height":
                        if (pair.Length > 1)
                            height = Misc.ParseLength(pair[1]);
                        break;
                    case "scale":
                        if (pair.Length > 1)
                            float.TryParse(pair[1], out scale);
                        break;
                    default:
                        break;
                }
            }

            // resize image according to optinal params

            if (width != 0 && height != 0)  // width and height was set in optinal parameters
            {
                imageShape.LockAspectRatio = MsoTriState.msoFalse;
                imageShape.Height = height;
                imageShape.Width = width;
            }
            else if (width != 0)  // only width was set
            {
                imageShape.LockAspectRatio = MsoTriState.msoTrue;
                imageShape.Width = width;
            }
            else if (height != 0)   // only height was set
            {
                imageShape.LockAspectRatio = MsoTriState.msoTrue;
                imageShape.Height = height;
            }
            else if (scale != 0) // scale was set
            {
                imageShape.LockAspectRatio = MsoTriState.msoTrue;
                imageShape.Width *= scale;
            }
        }

        /// <summary>
        /// Generate list
        /// </summary>
        /// <param name="items">List of items</param>
        /// <param name="shape">Text shape</param>
        /// <param name="level">Indentation level</param>
        /// <param name="type">Type of list (bulleted/numbered)</param>
        /// <returns>true if completed; false if paused</returns>
        private bool GenerateList(List<Node> items, PowerPoint.Shape shape, int level, MsoBulletType type)
        {
            Stack<Node> nodes = new Stack<Node>();

            Node currentNode;

            // skip expanding of child nodes to stack
            bool skip;

            // show bullet in item
            bool show_bullet;

            int itemStartAt, realItemStartAt;
            int itemsCount = 0;
            int paragraphs;

            foreach (Node item in items)
            {
                show_bullet = true;
                realItemStartAt = itemStartAt = shape.TextFrame2.TextRange.Text.Length;
                paragraphs = 0;
                itemsCount++;

                // copy content to stack
                foreach (Node node in item.Children.Reverse<Node>())
                {
                    nodes.Push(node);
                }

                // process nodes on stack
                while (nodes.Count != 0)
                {
                    currentNode = nodes.Pop();
                    skip = false;

                    // process node depending on its type
                    switch (currentNode.Type)
                    {
                        case "string":
                            paragraphs += (int) Misc.CountLinesInString(currentNode.Content as string);

                            _format.AppendText(shape, currentNode.Content as string);
                            break;
                        case "paragraph":
                            paragraphs++;
                            _format.AppendText(shape, "\r");
                            break;
                        case "pause":
                            if (Pause())
                            {
                                FormatListItem(shape, type, itemStartAt, level, paragraphs, itemsCount, show_bullet);
                                return false;
                            }
                            break;
                        case "today":
                            shape.TextFrame.TextRange.InsertDateTime(PowerPoint.PpDateTimeFormat.ppDateTimeFigureOut, MsoTriState.msoTrue);
                            break;
                        case "descriptionlist":
                        case "table":
                        case "image":
                            if (!Settings.Instance.NestedAsText)
                            {
                                skip = true;
                                _postProcessing.Enqueue(currentNode);
                            }
                            break;
                        case "bulletlist":
                        case "numberedlist":
                            skip = true;

                            FormatListItem(shape, type, itemStartAt, level, paragraphs, itemsCount, show_bullet);

                            _format.AppendText(shape, "\r");

                            if (!GenerateList(currentNode.Children, shape, level + 1, currentNode.Type == "bulletlist" ? MsoBulletType.msoBulletUnnumbered : MsoBulletType.msoBulletNumbered))
                                return false;

                            paragraphs = 0;
                            itemStartAt = shape.TextFrame2.TextRange.Text.Length;
                            show_bullet = false;

                            break;
                        default: // other -> check for simple formats
                            SimpleTextFormat(nodes, currentNode);
                            break;
                    }

                    if (currentNode.Children == null || skip)
                        continue;

                    // push child nodes to stack
                    foreach (Node node in currentNode.Children.Reverse<Node>())
                    {
                        nodes.Push(node);
                    }
                }

                FormatListItem(shape, type, itemStartAt, level, paragraphs, itemsCount, show_bullet);

                int min = item.OverlayList.Count != 0 ? item.OverlayList.Min() : int.MaxValue;
                _maxPass = Math.Max(Misc.MaxOverlay(item.OverlayList), _maxPass);    // set maximal number of passes from overlay specification
                if (!(item.OverlayList.Count == 0 || item.OverlayList.Contains(_passNumber) || min < 0 && Math.Abs(min) < _passNumber))
                {
                    shape.TextFrame2.TextRange.Characters[realItemStartAt + 1, shape.TextFrame2.TextRange.Text.Length - realItemStartAt - 1].Paragraphs.Font.Fill.Visible = MsoTriState.msoFalse;
                    shape.TextFrame2.TextRange.Characters[realItemStartAt + 1, shape.TextFrame2.TextRange.Text.Length - realItemStartAt - 1].Paragraphs.ParagraphFormat.Bullet.Visible = MsoTriState.msoFalse;
                }

                _format.AppendText(shape, "\r");
            }

            return true;
        }

        /// <summary>
        /// Setup list item (even partial) type, indent etc.
        /// </summary>
        /// <param name="shape">List text shape</param>
        /// <param name="type">Type of bullets</param>
        /// <param name="itemStartAt">Item starts at character</param>
        /// <param name="level">Indent level</param>
        /// <param name="paragraphs">Number of paragraphs</param>
        /// <param name="itemsCount">Number of items</param>
        /// <param name="show_bullet">Show bullet</param>
        private void FormatListItem(PowerPoint.Shape shape, MsoBulletType type, int itemStartAt, int level, int paragraphs, int itemsCount, bool show_bullet)
        {
            if (paragraphs == 0)
            {
                shape.TextFrame2.TextRange.Characters[itemStartAt + 1, shape.TextFrame2.TextRange.Text.Length - itemStartAt - 1].ParagraphFormat.Bullet.Type = type;
                shape.TextFrame2.TextRange.Characters[itemStartAt + 1, shape.TextFrame2.TextRange.Text.Length - itemStartAt - 1].ParagraphFormat.IndentLevel = level + 1;
                if (!show_bullet)
                {
                    shape.TextFrame2.TextRange.Characters[itemStartAt + 1, shape.TextFrame2.TextRange.Text.Length - itemStartAt - 1].ParagraphFormat.Bullet.Visible = MsoTriState.msoFalse;
                    shape.TextFrame2.TextRange.Characters[itemStartAt + 1, shape.TextFrame2.TextRange.Text.Length - itemStartAt - 1].ParagraphFormat.IndentLevel = level + 2;
                }
                if (type == MsoBulletType.msoBulletNumbered && show_bullet)
                    shape.TextFrame2.TextRange.Characters[itemStartAt + 1, shape.TextFrame2.TextRange.Text.Length - itemStartAt - 1].ParagraphFormat.Bullet.StartValue = itemsCount;
            }
            else
            {
                if (show_bullet)
                {
                    shape.TextFrame2.TextRange.Characters[itemStartAt + 1, shape.TextFrame2.TextRange.Text.Length - itemStartAt - 1].Paragraphs[1, 1].ParagraphFormat.Bullet.Type = type;
                    shape.TextFrame2.TextRange.Characters[itemStartAt + 1, shape.TextFrame2.TextRange.Text.Length - itemStartAt - 1].Paragraphs[1, 1].ParagraphFormat.IndentLevel = level + 1;
                    if (type == MsoBulletType.msoBulletNumbered)
                        shape.TextFrame2.TextRange.Characters[itemStartAt + 1, shape.TextFrame2.TextRange.Text.Length - itemStartAt - 1].Paragraphs[1, 1].ParagraphFormat.Bullet.StartValue = itemsCount;
                    shape.TextFrame2.TextRange.Characters[itemStartAt + 1, shape.TextFrame2.TextRange.Text.Length - itemStartAt - 1].Paragraphs[2, paragraphs].ParagraphFormat.Bullet.Type = type;
                    shape.TextFrame2.TextRange.Characters[itemStartAt + 1, shape.TextFrame2.TextRange.Text.Length - itemStartAt - 1].Paragraphs[2, paragraphs].ParagraphFormat.Bullet.Visible = MsoTriState.msoFalse;
                    shape.TextFrame2.TextRange.Characters[itemStartAt + 1, shape.TextFrame2.TextRange.Text.Length - itemStartAt - 1].Paragraphs[2, paragraphs].ParagraphFormat.IndentLevel = level + 2;
                }
                else
                {
                    shape.TextFrame2.TextRange.Characters[itemStartAt + 1, shape.TextFrame2.TextRange.Text.Length - itemStartAt - 1].Paragraphs.ParagraphFormat.Bullet.Type = type;
                    shape.TextFrame2.TextRange.Characters[itemStartAt + 1, shape.TextFrame2.TextRange.Text.Length - itemStartAt - 1].Paragraphs.ParagraphFormat.Bullet.Visible = MsoTriState.msoFalse;
                    shape.TextFrame2.TextRange.Characters[itemStartAt + 1, shape.TextFrame2.TextRange.Text.Length - itemStartAt - 1].Paragraphs.ParagraphFormat.IndentLevel = level + 2;
                }
            }
        }

        /// <summary>
        /// Generate description list
        /// </summary>
        /// <param name="items">List of items</param>
        /// <returns>true if completed; false if paused</returns>
        private bool GenerateDescriptionList(List<Node> items)
        {
            // create table shape with one row and two columns
            PowerPoint.Shape shape = _slide.Shapes.AddTable(1, 2, 36.0f, _bottomShapeBorder + 5.0f, 648.0f);
            // style without background and borders
            shape.Table.ApplyStyle("2D5ABB26-0587-4C30-8999-92F81FD0307C");

            // set columns width in ratio 1:2
            shape.Table.Columns[1].Width = 166;
            shape.Table.Columns[2].Width = 482;

            Stack<Node> nodes = new Stack<Node>();

            Node currentNode;

            int itemsCounter = 0;

            // how deep we are in nested lists
            int nested = 0;

            // skip expanding child nodes
            bool skip = false;

            PowerPoint.Shape termShape = null, definitionShape = null;

            foreach (Node node in items)
            {
                if (node.Type == "item")
                {
                    foreach (Node item in node.Children.Reverse<Node>())
                    {
                        nodes.Push(item);
                    }

                    if (nested == 0)
                    {
                        itemsCounter++;

                        // get cell shape
                        termShape = shape.Table.Cell(itemsCounter, 1).Shape;
                        definitionShape = shape.Table.Cell(itemsCounter, 2).Shape;

                        // set term
                        termShape.TextFrame2.TextRange.InsertAfter(node.OptionalParams as string);
                        termShape.TextFrame2.TextRange.Font.Bold = MsoTriState.msoTrue;
                        termShape.TextFrame2.TextRange.Font.Size = _baseFontSize * 2.0f;
                    }
                    // set definition
                    _format.Invalidate();

                    // process nodes on stack
                    while (nodes.Count != 0)
                    {
                        skip = false;

                        currentNode = nodes.Pop();

                        // process node depending on its type
                        switch (currentNode.Type)
                        {
                            case "string":
                                _format.AppendText(definitionShape, currentNode.Content as string);
                                break;
                            case "paragraph":
                                _format.AppendText(definitionShape, "\r");
                                break;
                            case "pause":
                                if (Pause())
                                    return false;
                                break;
                            case "image":
                                if (!Settings.Instance.NestedAsText)
                                {
                                    skip = true;
                                    _postProcessing.Enqueue(currentNode);
                                }
                                break;
                            case "table":
                                if (!Settings.Instance.NestedAsText)
                                {
                                    skip = true;
                                    _postProcessing.Enqueue(currentNode);
                                }
                                break;
                            case "descriptionlist":
                            case "bulletlist":
                            case "numberedlist":
                                if (Settings.Instance.NestedAsText)
                                {
                                    // increment nested counter and add __end_list node
                                    nested++;
                                    nodes.Push(new Node("__end_list"));
                                }
                                else
                                {
                                    // add node to post processing queue
                                    skip = true;
                                    _postProcessing.Enqueue(currentNode);
                                }
                                break;
                            case "__end_list":
                                nested--;
                                break;
                            case "today":
                                definitionShape.TextFrame.TextRange.InsertDateTime(PowerPoint.PpDateTimeFormat.ppDateTimeFigureOut, MsoTriState.msoTrue);
                                break;
                            default: // other -> check for simple formats
                                SimpleTextFormat(nodes, currentNode);
                                break;
                        }

                        if (currentNode.Children == null || skip)
                            continue;

                        // push child nodes to stack
                        foreach (Node item in currentNode.Children.Reverse<Node>())
                        {
                            nodes.Push(item);
                        }
                    }

                    if (nested == 0)
                    {
                        // add row to the end of table
                        shape.Table.Rows.Add();

                        // check overlays
                        int min = node.OverlayList.Count != 0 ? node.OverlayList.Min() : int.MaxValue;
                        _maxPass = Math.Max(Misc.MaxOverlay(node.OverlayList), _maxPass);    // set maximal number of passes from overlay specification
                        if (!(node.OverlayList.Count == 0 || node.OverlayList.Contains(_passNumber) || min < 0 && Math.Abs(min) < _passNumber))
                        {
                            termShape.TextFrame2.TextRange.Font.Fill.Visible = MsoTriState.msoFalse;
                            definitionShape.TextFrame2.TextRange.Font.Fill.Visible = MsoTriState.msoFalse;
                        }
                    }
                }
            }

            // delete last row
            shape.Table.Rows[shape.Table.Rows.Count].Delete();

            return true;
        }

        /// <summary>
        /// Generate title page
        /// </summary>
        private void GenerateTitlePage()
        {
            if (_titlesettings.Count == 0)  // no title settings -> don't generate shape
                return;

            PowerPoint.Shape shape = _slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 36.0f, _bottomShapeBorder + 50.0f, 648.0f, 10.0f);
            // set alignment to center
            shape.TextFrame2.TextRange.ParagraphFormat.Alignment = MsoParagraphAlignment.msoAlignCenter;

            string[] keys = { "title", "author", "date" };

            Stack<Node> nodes = new Stack<Node>();
            Node currentNode;

            foreach (string key in keys)
            {
                if (_titlesettings.ContainsKey(key))
                {
                    if (key == "title")
                        SimpleTextFormat(nodes, new Node("huge"));

                    if (key == "author")
                        SimpleTextFormat(nodes, new Node("large"));

                    // copy column content to stack
                    foreach (Node item in _titlesettings[key].Reverse<Node>())
                    {
                        nodes.Push(item);
                    }

                    // process nodes on stack
                    while (nodes.Count != 0)
                    {
                        currentNode = nodes.Pop();

                        // process node depending on its type
                        switch (currentNode.Type)
                        {
                            case "string":
                                _format.AppendText(shape, currentNode.Content as string);
                                break;
                            case "paragraph":
                                _format.AppendText(shape, "\r");
                                break;
                            case "today":
                                shape.TextFrame.TextRange.InsertDateTime(PowerPoint.PpDateTimeFormat.ppDateTimeFigureOut, MsoTriState.msoTrue);
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

                    _format.Invalidate();

                    _format.AppendText(shape, "\r\r\r\r");
                }
            }
        }

        /// <summary>
        /// Post processing - generate shapes (elements) from queue
        /// </summary>
        private void PostProcessing()
        {
            Node node;

            while (_postProcessing.Count > 0)
            {
                node = _postProcessing.Dequeue();

                switch (node.Type)
                {
                    case "image":
                        UpdateBottomShapeBorder(false);
                        PowerPoint.Shape imageShape;

                        GenerateImage(node, out imageShape);
                        break;
                    case "bulletlist":
                    case "numberedlist":
                        UpdateBottomShapeBorder(false);
                        PowerPoint.Shape shape = _slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 36.0f, _bottomShapeBorder + 5.0f, 648.0f, 10.0f);

                        GenerateList(node.Children, shape, 1, node.Type == "bulletlist" ? MsoBulletType.msoBulletUnnumbered : MsoBulletType.msoBulletNumbered);
                        break;
                    case "table":
                        UpdateBottomShapeBorder(false);
                        PowerPoint.Shape tableShape;

                        GenerateTable(node, out tableShape);
                        break;
                    case "descriptionlist":
                        UpdateBottomShapeBorder(false);

                        GenerateDescriptionList(node.Children);
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Reposition shapes on slide
        /// </summary>
        /// <param name="shapes">List of shapes to reposition (must contain two shapes for repositioning)</param>
        private void Reshaper(List<PowerPoint.Shape> shapes)
        {
            while (shapes.Remove(null)) // remove all null shapes from list
                ;

            PowerPoint.Shape lastMovableShape;

            if (shapes.Count == 2)
            {
                lastMovableShape = shapes[1];

                // text shape + image or table
                if(shapes[0].HasTextFrame == MsoTriState.msoTrue && shapes[1].HasTextFrame == MsoTriState.msoFalse) {

                    // check if second shape can possibly fit after first one
                    if (shapes[0].Left + shapes[0].TextFrame2.TextRange.Lines[shapes[0].TextFrame2.TextRange.Lines.Count].BoundWidth + shapes[0].TextFrame2.MarginLeft + shapes[0].TextFrame2.MarginRight + 5.0f + shapes[1].Width <= 648.0f + 36.0f)
                    {
                        PowerPoint.Shape lastLineShape;

                        // remove last line from text shape and move it to the new shape
                        if (shapes[0].TextFrame2.TextRange.Lines.Count > 1)
                        {
                            shapes[0].TextFrame2.TextRange.Lines[shapes[0].TextFrame2.TextRange.Lines.Count].Cut();
                            lastLineShape = _slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, shapes[0].Left, shapes[0].Top + shapes[0].Height + 5.0f, 648.0f, 10.0f);
                            lastLineShape.TextFrame2.TextRange.Paste();
                            lastMovableShape = lastLineShape;
                        }
                        else
                        {
                            lastLineShape = shapes[0];
                        }

                        // change shape width according ot its bounding box
                        lastLineShape.Width = lastLineShape.TextFrame2.TextRange.BoundWidth + lastLineShape.TextFrame2.MarginRight + lastLineShape.TextFrame2.MarginLeft + 1;

                        shapes[1].Top = lastLineShape.Top;
                        shapes[1].Left = lastLineShape.Left + lastLineShape.Width + 5.0f;
                    }

                    shapes.Clear();
                    shapes.Add(lastMovableShape);
                }
                // image or table + image or table
                else if (shapes[0].HasTextFrame == MsoTriState.msoFalse && shapes[1].HasTextFrame == MsoTriState.msoFalse)
                {
                    // check if second shape can possibly fit after first one
                    if (shapes[0].Left + shapes[0].Width + 5.0f + shapes[1].Width <= 648.0f + 36.0f)
                    {
                        shapes[1].Top = shapes[0].Top;
                        shapes[1].Left = shapes[0].Left + shapes[0].Width + 5.0f;
                    }

                    shapes.Clear();
                    shapes.Add(lastMovableShape);
                }
                // image or table + text shape
                else if (shapes[0].HasTextFrame == MsoTriState.msoFalse && shapes[1].HasTextFrame == MsoTriState.msoTrue)
                {
                    // get bounding box of first word
                    float wordBoundWidth = shapes[1].TextFrame2.TextRange.Words[1, 1].BoundWidth + shapes[1].TextFrame2.MarginLeft + shapes[1].TextFrame2.MarginRight;

                    // check if at least one word from textbox will fit after first shape
                    if (shapes[0].Left + shapes[0].Width + 5.0f + wordBoundWidth <= 648.0f + 36.0f)
                    {
                        shapes[1].TextFrame2.AutoSize = MsoAutoSize.msoAutoSizeShapeToFitText;
                        shapes[1].Top = shapes[0].Top;
                        shapes[1].Left = shapes[0].Left + shapes[0].Width + 5.0f;
                        shapes[1].Width = 648.0f + 36.0f - shapes[1].Left;

                        if (shapes[1].TextFrame2.TextRange.Lines.Count > 1)
                        {
                            shapes[1].TextFrame2.TextRange.Lines[2, shapes[1].TextFrame2.TextRange.Lines.Count - 1].Cut();
                            PowerPoint.Shape lastLinesShape = _slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, shapes[0].Left, shapes[0].Top + shapes[0].Height + 5.0f, 648.0f, 10.0f);
                            lastLinesShape.TextFrame2.TextRange.Paste();
                            lastMovableShape = lastLinesShape;
                        }
                    }

                    shapes.Clear();
                    shapes.Add(lastMovableShape);
                }
                else    // text shape + text shape (shouldn't happen)
                {
                    shapes.RemoveRange(0, 2);   // don't reshape and remove them from list
                }
            }
        }
    }
}
