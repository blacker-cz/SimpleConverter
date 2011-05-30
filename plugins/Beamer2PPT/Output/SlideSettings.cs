using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    /// <summary>
    /// Class containing slide settings
    /// </summary>
    class SlideSettings
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="parameters">Slide parameters</param>
        public SlideSettings(string parameters)
        {
            ContentAlign = Align.CENTER;

            ParseParams(parameters);
        }

        /// <summary>
        /// Content vertical align
        /// </summary>
        public Align ContentAlign { get; private set; }

        /// <summary>
        /// Parse slide parameters
        /// </summary>
        /// <param name="parameters"></param>
        private void ParseParams(string parameters)
        {
            if (parameters == null || parameters.Trim().Length == 0)
                return;

            // "shrink,squeeze,c" -> ["shrink","squeeze","c"]
            string[] parts = parameters.Replace(" ", "").Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

            foreach (string part in parts)
            {
                switch (part)
                {
                    case "t":
                        ContentAlign = Align.TOP;
                        break;
                    case "c":
                        ContentAlign = Align.CENTER;
                        break;
                    case "b":
                        ContentAlign = Align.BOTTOM;
                        break;
                    default:
                        break;
                }
            }
        }

        /// <summary>
        /// Content vertical align types
        /// </summary>
        public enum Align
        {
            TOP, CENTER, BOTTOM
        }
    }
}
