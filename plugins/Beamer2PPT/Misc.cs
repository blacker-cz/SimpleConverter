using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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

            if (overlaySpecification.Trim().Length == 0)    // empty overlay specification -> return empty set
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
    }
}
