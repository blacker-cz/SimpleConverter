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
    /// Text formatting class
    /// </summary>
    class TextFormat
    {
        /// <summary>
        /// Base font size.
        /// </summary>
        private float _baseFontSize;

        /// <summary>
        /// Stack of format settings
        /// </summary>
        private Stack<FormatSettings> _settingsStack;

        /// <summary>
        /// Current format settings
        /// </summary>
        private FormatSettings _currentSettings;

        /// <summary>
        /// Information if format has changed since last append
        /// </summary>
        private bool _changed;

        /// <summary>
        /// Information if format was applied before
        /// </summary>
        private bool _firstRun;

        /// <summary>
        /// Default parameterless constructor.
        /// Base font size is set to 11pt
        /// </summary>
        public TextFormat() : this (11.0f)
        {
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="baseFontSize">Base font size (11pt, 12pt, ...)</param>
        /// <param name="isTitle">Flag if shape is title shape (uses different default font) - optional, default: false</param>
        public TextFormat(float baseFontSize, bool isTitle = false)
        {
            _baseFontSize = baseFontSize * 2.0f;

            _settingsStack = new Stack<FormatSettings>();

            _currentSettings = new FormatSettings(_baseFontSize, isTitle);

            _changed = true;

            _firstRun = true;
        }

        /// <summary>
        /// Modify current formatting according to node.
        /// </summary>
        /// <param name="node">Node containing new font settings</param>
        public void ModifyFormat(Node node)
        {
            // save current font settings
            _settingsStack.Push(_currentSettings.Clone() as FormatSettings);

            _changed = true;

            switch (node.Type)
            {
                case "bold":
                    _currentSettings.Bold = MsoTriState.msoTrue;
                    break;
                case "italic":
                    _currentSettings.Italic = MsoTriState.msoTrue;
                    break;
                case "underline":
                    _currentSettings.Underline = MsoTextUnderlineType.msoUnderlineSingleLine;
                    break;
                case "smallcaps":
                    _currentSettings.Smallcaps = MsoTriState.msoTrue;
                    break;
                case "typewriter":
                    _currentSettings.FontFamily = @"Courier New";
                    break;
                case "color":
                    _currentSettings.Color = ParseColor(node.OptionalParams, node.Content as string);
                    break;
                // coeficients for relative font size are computed from default font size for each LaTeX size command
                case "tiny":
                    _currentSettings.FontSize = _baseFontSize / 2f;
                    break;
                case "scriptsize":
                    _currentSettings.FontSize = _baseFontSize / 1.4285714285714285714285714285714f;
                    break;
                case "footnotesize":
                    _currentSettings.FontSize = _baseFontSize / 1.25f;
                    break;
                case "small":
                    _currentSettings.FontSize = _baseFontSize / 1.1111111111111111111111111111111f;
                    break;
                case "normalsize":
                    _currentSettings.FontSize = _baseFontSize;
                    break;
                case "large":
                    _currentSettings.FontSize = _baseFontSize * 1.2f;
                    break;
                case "Large":
                    _currentSettings.FontSize = _baseFontSize * 1.44f;
                    break;
                case "LARGE":
                    _currentSettings.FontSize = _baseFontSize * 1.728f;
                    break;
                case "huge":
                    _currentSettings.FontSize = _baseFontSize * 2.0736f;
                    break;
                case "Huge":
                    _currentSettings.FontSize = _baseFontSize * 2.48832f;
                    break;
                default:
                    _settingsStack.Pop();   // no changes -> throw away stack top (saved at start of this method)
                    _changed = false;
                    break;
            }
        }

        /// <summary>
        /// Return current formatting to previous state.
        /// </summary>
        public void RollBackFormat()
        {
            if(_settingsStack.Count > 0)
                _currentSettings = _settingsStack.Pop();

            _changed = true;
        }

        /// <summary>
        /// Force format update during next append
        /// </summary>
        public void Invalidate()
        {
            _changed = true;
        }

        /// <summary>
        /// Append text with current internal formatting to the end of shape.
        /// </summary>
        /// <param name="shape">Text shape</param>
        /// <param name="text">Appended text</param>
        public void AppendText(PowerPoint.Shape shape, string text)
        {
            if (shape.HasTextFrame != MsoTriState.msoTrue)
                throw new ArgumentException("Shape must contain text frame.");

            AppendText(shape.TextFrame2.TextRange, text);
        }

        /// <summary>
        /// Append text with current internal formatting to the end of text range.
        /// </summary>
        /// <param name="range">Text range</param>
        /// <param name="text">Appended text</param>
        public void AppendText(TextRange2 range, string text)
        {
            int start = range.Text.Length;

            text = text.Replace("\r\n", "\r");
            text = text.Replace("\n", "\r");

            // filter spaces (no space at beginning of line, and no space after space)
            if (text.StartsWith(" ") && (range.Text.EndsWith(" ") || range.Text.EndsWith("\u00A0") || range.Text.Length == 0 || range.Text.EndsWith("\r")))
                text = text.TrimStart(' ');

            if (text.Length > 0)
                range.InsertAfter(text);
            else
                return;

            // apply formatting only if there were changes (experimental!)
            if (_changed)
            {
                TextRange2 format = range.Characters[start + 1, text.Length];

                // check used color, if color is null set default theme color
                if (_currentSettings.Color != null)
                    format.Font.Fill.ForeColor.RGB = (int) _currentSettings.Color;
                else
                    format.Font.Fill.ForeColor.ObjectThemeColor = MsoThemeColorIndex.msoThemeColorText1;

                format.Font.Bold = _currentSettings.Bold;
                format.Font.Italic = _currentSettings.Italic;
                format.Font.Name = _currentSettings.FontFamily;
                format.Font.UnderlineStyle = _currentSettings.Underline;
                format.Font.Smallcaps = _currentSettings.Smallcaps;
                format.Font.Size = _currentSettings.FontSize;
                _changed = false;
            }
        }

        /// <summary>
        /// Get color number from string
        /// </summary>
        /// <param name="optional">Optional color settings</param>
        /// <param name="color">Color string</param>
        /// <returns>Color acceptable by PowerPoint</returns>
        private int ParseColor(string optional, string color)
        {
            Regex regex;
            Match match;
            int retval = 0;

            if (optional != null && optional.Length > 0 && color != null && color.Length > 0)
            {
                switch (optional.Trim())
                {
                    case "gray":
                        regex = new Regex(@"^([0-1](\.[0-9]*)?)$", RegexOptions.IgnoreCase);
                        match = regex.Match(color.Trim());
                        
                        if (match.Success)
                        {
                            float degree;

                            if (!float.TryParse(match.Groups[1].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out degree))
                                return 0;

                            degree = Math.Min(degree, 1.0f);

                            retval += (int)(degree * 0xFF);
                            retval <<= 8;   // shift 8 bits
                            retval += (int)(degree * 0xFF);
                            retval <<= 8;   // shift 8 bits
                            retval += (int)(degree * 0xFF);

                            return retval;
                        }

                        return 0;
                    case "rgb":
                        regex = new Regex(@"^([0-1](\.[0-9]*)?),([0-1](\.[0-9]*)?),([0-1](\.[0-9]*)?)$", RegexOptions.IgnoreCase);
                        match = regex.Match(color.Trim());

                        if (match.Success)
                        {
                            float r, g, b;

                            if (!float.TryParse(match.Groups[1].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out r))
                                return 0;
                            if (!float.TryParse(match.Groups[3].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out g))
                                return 0;
                            if (!float.TryParse(match.Groups[5].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out b))
                                return 0;

                            r = Math.Min(r, 1.0f);
                            g = Math.Min(g, 1.0f);
                            b = Math.Min(b, 1.0f);

                            retval += (int)(b * 0xFF);
                            retval <<= 8;   // shift 8 bits
                            retval += (int)(g * 0xFF);
                            retval <<= 8;   // shift 8 bits
                            retval += (int)(r * 0xFF);

                            return retval;
                        }

                        return 0;
                    case "RGB":
                        regex = new Regex(@"^([0-9]{1,3}),([0-9]{1,3}),([0-9]{1,3})$", RegexOptions.IgnoreCase);
                        match = regex.Match(color.Trim());

                        if (match.Success)
                        {
                            int r, g, b;

                            if (!int.TryParse(match.Groups[1].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out r))
                                return 0;
                            if (!int.TryParse(match.Groups[2].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out g))
                                return 0;
                            if (!int.TryParse(match.Groups[3].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out b))
                                return 0;

                            r = Math.Min(r, 255);
                            g = Math.Min(g, 255);
                            b = Math.Min(b, 255);

                            retval += b;
                            retval <<= 8;   // shift 8 bits
                            retval += g;
                            retval <<= 8;   // shift 8 bits
                            retval += r;

                            return retval;
                        }

                        return 0;
                    case "HTML":
                        regex = new Regex(@"^([A-F0-9]{2})([A-F0-9]{2})([A-F0-9]{2})$", RegexOptions.IgnoreCase);
                        match = regex.Match(color.Trim());
                        
                        if (match.Success)
                        {
                            string hex = match.Groups[3].Value + match.Groups[2].Value + match.Groups[1].Value;

                            if (!int.TryParse(hex, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out retval))
                                return 0;
                            
                            return retval;
                        }

                        return 0;
                    default:
                        break;
                }
            }

            if (color != null && color.Length > 0)
            {
                switch (color.Trim().ToLower())
                {
                    case "white":
                        return 0xFFFFFF;
                    case "black":
                        return 0;
                    case "red":
                        return 0x0000FF;
                    case "green":
                        return 0x00FF00;
                    case "blue":
                        return 0xFF0000;
                    case "cyan":
                        return 0xFFFF00;
                    case "magenta":
                        return 0xFF00FF;
                    case "yellow":
                        return 0X00FFFF;
                    default:
                        break;
                }
            }

            return 0;   // not found return black color
        }

        /// <summary>
        /// Private class used for stacking format settings
        /// </summary>
        private class FormatSettings : ICloneable
        {
            #region Public properties

            /// <summary>
            /// Font family
            /// </summary>
            public string FontFamily { get; set; }

            /// <summary>
            /// Font size
            /// </summary>
            public float FontSize { get; set; }

            /// <summary>
            /// Font color
            /// </summary>
            public int? Color { get; set; }

            /// <summary>
            /// Bold
            /// </summary>
            public MsoTriState Bold { get; set; }

            /// <summary>
            /// Italic
            /// </summary>
            public MsoTriState Italic { get; set; }

            /// <summary>
            /// Underline
            /// </summary>
            public MsoTextUnderlineType Underline { get; set; }

            /// <summary>
            /// Smallcaps
            /// </summary>
            public MsoTriState Smallcaps { get; set; }

            #endregion // Public properties

            #region Constructors

            /// <summary>
            /// Default constructor
            /// </summary>
            /// <param name="fontSize">Font size</param>
            /// <param name="isTitle">Flag if shape is title shape (uses different default font)</param>
            public FormatSettings(float fontSize, bool isTitle)
            {
                // To assign a Headings (major) or Body (minor) font style to text, you change the font name to this:
                //
                // "+" & FontType & "-" & FontLang
                //
                // FontType:
                //
                //    Major (Headings) = "mj"
                //    Minor (Body) = "mn"
                //
                // FontLang:
                //
                //    Latin = "lt"
                //    Complex Scripts = "cs"
                //    East Asian = "ea"
                // (source: http://pptfaq.com/FAQ00957.htm)

                if(isTitle)
                    FontFamily = @"+mj-lt";
                else
                    FontFamily = @"+mn-lt";

                FontSize = fontSize;
                Color = null;  // default black color
                Bold = MsoTriState.msoFalse;
                Italic = MsoTriState.msoFalse;
                Underline = MsoTextUnderlineType.msoNoUnderline;
                Smallcaps = MsoTriState.msoFalse;
            }

            /// <summary>
            /// Private copy constructor
            /// </summary>
            /// <param name="format">Format settings object</param>
            private FormatSettings(FormatSettings format)
            {
                FontFamily = format.FontFamily;
                FontSize = format.FontSize;
                Color = format.Color;
                Bold = format.Bold;
                Italic = format.Italic;
                Underline = format.Underline;
                Smallcaps = format.Smallcaps;
            }

            #endregion // Constructors

            #region Implementation of ICloneable

            /// <summary>
            /// Clone object
            /// </summary>
            /// <returns>Clone of current object</returns>
            public object Clone()
            {
                return new FormatSettings(this);
            }

            #endregion // Implementation of ICloneable
        }
    }
}
