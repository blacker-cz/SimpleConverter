using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    /// <summary>
    /// Class holding tabular environment settings
    /// </summary>
    class TabularSettings
    {
        /// <summary>
        /// List of columns
        /// </summary>
        public List<Column> Columns { get; private set; }

        /// <summary>
        /// Set of borders between columns.
        /// Borders are indexed from 0 (0 is leftmost one, Columns.Count is the rightmost one).
        /// </summary>
        /// <example>
        /// string representation: c|c|c c|
        /// parsed representation: 1,2,4
        /// </example>
        public ISet<int> Borders { get; private set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public TabularSettings()
        {
            Columns = new List<Column>();
            Borders = new HashSet<int>();
        }

        /// <summary>
        /// Public constructor
        /// </summary>
        /// <param name="tabularSettings"></param>
        public TabularSettings(string tabularSettings) : this()
        {
            ParseHeader(tabularSettings);
        }

        /// <summary>
        /// Static parse method
        /// </summary>
        /// <param name="tabularSettings"></param>
        /// <returns>Instance of tabular settings class</returns>
        public static TabularSettings Parse(string tabularSettings)
        {
            return new TabularSettings(tabularSettings);
        }

        /// <summary>
        /// Parse tabular header
        /// </summary>
        /// <param name="tabularSettings">String with tabular settings</param>
        /// <returns>Instance of TabularSettings class</returns>
        private void ParseHeader(string tabularSettings)
        {
            // remove spaces
            tabularSettings = tabularSettings.Replace(" ", "");
            
            // split params
            string[] parts = tabularSettings.Split('{', '}');

            bool normal = true;
            bool width = false;

            int columnIndex = 0;

            foreach (string part in parts)
            {
                if (normal)
                {
                    foreach (char ch in part)
                    {
                        switch (ch)
                        {
                            case '|':
                                Borders.Add(columnIndex);
                                break;
                            case 'c':
                            case 'l':
                            case 'r':
                                columnIndex++;
                                Columns.Add(new Column(ch));
                                break;
                            case 'p':
                                normal = false;
                                width = true;
                                break;
                            default:
                                break;
                        }
                    }
                } else if (width)
                {
                    // column width here
                }
            }
        }

        /// <summary>
        /// Structure for keeping column alignment and optional width
        /// </summary>
        public struct Column
        {
            /// <summary>
            /// Constructor
            /// </summary>
            /// <param name="alignment">Columnt alignment</param>
            /// <param name="width">Columnt width (optional)</param>
            public Column(char alignment, float width = 0.0f)
            {
                this.alignment = alignment;
                this.width = width;
            }

            /// <summary>
            /// Column alignment
            /// </summary>
            public char alignment;

            /// <summary>
            /// Column width.
            /// Zero if not specified.
            /// </summary>
            public float width;
        }
    }
}
