using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Globalization;

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
    }
}
