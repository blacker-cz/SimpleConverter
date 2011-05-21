using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SimpleConverter
{
    /// <summary>
    /// Wraper class for files listed in files ListBox
    /// </summary>
    public class ListFile
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="filename">Complete file name (with path)</param>
        /// <param name="valid">Flag if file is valid for current plugin</param>
        public ListFile(string filename, bool valid = false)
        {
            Filename = System.IO.Path.GetFileName(filename); ;
            Filepath = filename;
            Valid = valid;
        }

        /// <summary>
        /// File name without path
        /// </summary>
        public string Filename { get; private set; }

        /// <summary>
        /// Complete file name (with path)
        /// </summary>
        public string Filepath { get; private set; }

        /// <summary>
        /// Flaf if file is valid for current plugin
        /// </summary>
        public bool Valid { get; set; }

        /// <summary>
        /// Background color of ListBox item (depends on <see cref="Valid"/>)
        /// </summary>
        public string ValidColor
        {
            get
            {
                if (Valid)
                    return "PaleGreen";
                else
                    return "LightSalmon";
            }
        }
    }
}
