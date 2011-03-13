using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SimpleConverter
{
    public class ListFile
    {
        public ListFile(string filename, bool valid = false)
        {
            Filename = System.IO.Path.GetFileName(filename); ;
            Filepath = filename;
            Valid = valid;
        }

        public string Filename { get; private set; }

        public string Filepath { get; private set; }

        public bool Valid { get; set; }

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
