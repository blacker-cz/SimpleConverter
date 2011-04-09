﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    /// <summary>
    /// Miscellaneous methods
    /// </summary>
    static class Misc
    {
        /// <summary>
        /// Method for parsing string overlay specification.
        /// </summary>
        /// <example>
        /// If overlay specification contains for example: 1,2-3,4-
        /// Then returned set will contain: 1,2,3,4,-4
        /// where -4 denotes that element will appear on pass 4 and greater
        /// </example>
        /// <param name="overlaySpecification">Overlay specification</param>
        /// <returns>Set of pass numbers</returns>
        public static ISet<int> ParseOverlay(string overlaySpecification)
        {
            ISet<int> overlays = new HashSet<int>();

            // empty overlay specification -> return empty set
            if (overlaySpecification == null || overlaySpecification.Trim().Length == 0)
                return overlays;

            // "1,2-3,4-" -> ["1","2-3","4-"]
            string[] parts = overlaySpecification.Trim().Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            int number;

            foreach (string part in parts)
            {
                try
                {
                    if (part.Contains('-'))
                    {
                        if (part.Trim().StartsWith("-"))
                        {
                            number = int.Parse(part.Trim().Substring(1));
                            if (number > 0)
                                overlays.Add(number);
                        }
                        else if (part.Trim().EndsWith("-"))
                        {
                            number = int.Parse(part.Trim().Substring(0, part.Trim().Length-1));
                            if (number > 0) {
                                overlays.Add(number);
                                overlays.Add(-number);
                            }
                        }
                        else    // '-' is between two numbers
                        {
                            string[] boundaries = part.Trim().Split(new char[] { '-' }, StringSplitOptions.RemoveEmptyEntries);

                            if (boundaries.Length != 2)
                                continue;

                            int a = int.Parse(boundaries[0]);
                            int b = int.Parse(boundaries[2]);

                            if (b < a)  // if second index is lesser than first -> set only "to infinity" flag (negative value)
                            {
                                overlays.Add(-a);
                                continue;
                            }

                            for (int i = a; i <= b; i++)
                            {
                                overlays.Add(i);
                            }
                        }
                    }
                    else
                    {
                        number = int.Parse(part);
                        overlays.Add(number);
                    }
                }
                catch { } // best effort
            }

            return overlays;
        }

        /// <summary>
        /// Get max overlay number
        /// </summary>
        /// <param name="overlays">Overlays list</param>
        /// <returns>Maximal overlay</returns>
        public static int MaxOverlay(ISet<int> overlays)
        {
            if (overlays.Count == 0)
                return 1;

            return Math.Max(overlays.Max(), Math.Abs(overlays.Min()));
        }

        /// <summary>
        /// Parse length string (cm, in, pt, mm)
        /// </summary>
        /// <param name="part">Length string</param>
        /// <returns>Length on slide</returns>
        public static float ParseLength(string part)
        {
            // note:
            //      1 pt = 1 pt
            //      1 mm = 2.84 pt
            //      1 cm = 28.4 pt
            //      1 in = 72.27 pt

            Regex regex = new Regex(@"^([0-9]+(\.[0-9]*)?) *(cm|in|pt|mm)$", RegexOptions.IgnoreCase);

            Match match = regex.Match(part.Trim());

            if (match.Success)
            {
                float length;

                if (!float.TryParse(match.Groups[1].Value, System.Globalization.NumberStyles.Float, CultureInfo.InvariantCulture, out length))
                    return 0.0f;

                switch (match.Groups[3].Value.ToLower())
                {
                    case "pt":
                        return length;
                    case "mm":
                        return length * 2.84f;
                    case "cm":
                        return length * 28.4f;
                    case "in":
                        return length * 72.27f;
                    default:
                        break;
                }
            }

            // no match
            return 0.0f;
        }

        /// <summary>
        /// Trim whitespace from end of shape content
        /// </summary>
        /// <param name="shape">Shape</param>
        public static void TrimShape(PowerPoint.Shape shape)
        {
            if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame2.HasText == MsoTriState.msoTrue)
            {
                // TextRange.TrimText() method is useless because it doesn't actually remove whitespaces from text range but returns its copy
                // first compute number of whitespace characters at the end of shape
                int size = shape.TextFrame2.TextRange.Text.Length - shape.TextFrame2.TextRange.Text.TrimEnd().Length;
                if (size > 0)   // then if there is more then zero of these characters -> delete them
                    shape.TextFrame2.TextRange.Characters[1 + shape.TextFrame2.TextRange.Text.Length - size, size].Delete();
            }
        }

        /// <summary>
        /// Update columns width to fit text (can only decrease size of columns!!)
        /// </summary>
        /// <param name="shape">Table shape</param>
        /// <param name="settings">Tabular settings (with information about columns)</param>
        /// <exception cref="ArgumentException"></exception>
        public static void AutoFitColumn(PowerPoint.Shape shape, TabularSettings settings)
        {
            if (shape.HasTable != MsoTriState.msoTrue)
                throw new ArgumentException("Shape must have table.");

            PowerPoint.TextFrame2 frame;

            float width;

            for (int column = 1; column <= shape.Table.Columns.Count; column++)
            {
                if (settings.Columns[column - 1].alignment != 'p' || (settings.Columns[column - 1].alignment == 'p' && settings.Columns[column - 1].width == 0))
                {
                    width = 0.0f;
                    for (int row = 1; row <= shape.Table.Rows.Count; row++)
                    {
                        frame = shape.Table.Cell(row, column).Shape.TextFrame2;

                        width = Math.Max(width, frame.TextRange.BoundWidth + frame.MarginLeft + frame.MarginRight + 1);
                    }
                }
                else
                {
                    width = settings.Columns[column - 1].width;
                }
                
                shape.Table.Columns[column].Width = width;
            }
        }
    }
}
