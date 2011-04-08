using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

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
        public TextFormat(float baseFontSize)
        {
            _baseFontSize = baseFontSize * 2.0f;

            _settingsStack = new Stack<FormatSettings>();

            _currentSettings = new FormatSettings(_baseFontSize);

            _changed = true;
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
            _currentSettings = _settingsStack.Pop();

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

            range.InsertAfter(text);

            // apply formatting only if there were changes (experimental!)
            if (_changed)
            {
                TextRange2 format = range.Characters[start + 1, text.Length];

                format.Font.Bold = _currentSettings.Bold;
                format.Font.Fill.ForeColor.RGB = _currentSettings.Color;
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
            throw new NotImplementedException();
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
            public int Color { get; set; }

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
            public FormatSettings(float fontSize)
            {
                FontFamily = @"Calibri";
                FontSize = fontSize;
                Color = 0;  // default black color
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
