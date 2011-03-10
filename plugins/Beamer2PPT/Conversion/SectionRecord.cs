using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    /// <summary>
    /// Section record class
    /// </summary>
    public class SectionRecord
    {
        /// <summary>
        /// Public constructor for section record
        /// </summary>
        /// <param name="slide">Slide number</param>
        /// <param name="content">Content of section record</param>
        /// <param name="type">Section level (type)</param>
        public SectionRecord(int slide, List<Node> content, SectionType type = SectionType.SECTION)
        {
            Type = type;
            Content = content;
            Slide = slide;
        }

        /// <summary>
        /// Section type (level)
        /// </summary>
        public SectionType Type { get; private set; }

        /// <summary>
        /// Section record content
        /// </summary>
        public List<Node> Content { get; private set; }

        /// <summary>
        /// Slide number
        /// </summary>
        public int Slide { get; private set; }
    }

    /// <summary>
    /// Allowed types of section records
    /// </summary>
    public enum SectionType
    {
        SECTION, SUBSECTION, SUBSUBSECTION
    }
}
