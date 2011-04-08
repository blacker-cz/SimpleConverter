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
        /// <exception cref="DocumentBuilderException"></exception>
        private void ParseHeader(string tabularSettings)
        {
            // remove spaces
            tabularSettings = tabularSettings.Replace(" ", "");
            
            // split params (with empty strings)
            string[] parts = tabularSettings.Split('{', '}');

            int state = 1;
            int iterations = 0;

            int columnIndex = 0;

            List<string> iterables = new List<string>();

            foreach (string part in parts)
            {
                switch (state)
                {
                    case 1:
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
                                    state = 6;
                                    break;
                                case '*':
                                    state = 2;
                                    break;
                                default:
                                    break;
                            }
                        }
                        break;

                    case 2: // number of iterations
                        if (!int.TryParse(part, out iterations))
                            throw new DocumentBuilderException("Invalid table column definition.");
                        state = 3;
                        break;

                    case 3: // empty string between } and { (after number of iterations)
                        if (part.Length != 0)
                            throw new DocumentBuilderException("Invalid table column definition.");
                        iterables.Clear();
                        state = 4;
                        break;

                    case 4: // iterated columns definition
                        iterables.Add(part);

                        if (part.EndsWith("p")) // following is width definition
                        {
                            state = 5;
                        }
                        else    // iterate through columns
                        {
                            for (int i = 0; i < iterations; i++)
                            {
                                int substate = 1;
                                foreach (string def in iterables)
                                {
                                    switch (substate)
                                    {
                                        case 1:
                                            foreach (char ch in def)
                                            {
                                                substate = 1;
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
                                                        substate = 2;
                                                        break;
                                                    case '*':
                                                        throw new DocumentBuilderException("Invalid table column definition.");
                                                    default:
                                                        break;
                                                }
                                            }
                                            break;

                                        case 2:
                                            columnIndex++;
                                            Columns.Add(new Column('p', Misc.ParseLength(def)));
                                            substate = 1;
                                            break;

                                        default:
                                            break;
                                    }
                                }
                            }
                            state = 1;
                        }
                        break;

                    case 5: // width definition inside iterated columns definition
                        iterables.Add(part);
                        state = 4;
                        break;

                    case 6: // column width
                        columnIndex++;
                        Columns.Add(new Column('p', Misc.ParseLength(part)));
                        state = 1;
                        break;

                    default:
                        break;
                }
            }

            return;
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
