using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    /// <summary>
    /// Frame title record class
    /// </summary>
    public class FrametitleRecord
    {
        /// <summary>
        /// Public constructor
        /// </summary>
        /// <param name="content">Content of frame title</param>
        /// <param name="content2">Content of frame subtitle</param>
        public FrametitleRecord(List<Node> title, List<Node> subtitle)
        {
            Title = title;
            Subtitle = subtitle;
        }

        /// <summary>
        /// Frame title content
        /// </summary>
        public List<Node> Title { get; set; }

        /// <summary>
        /// Frame subtitle content
        /// </summary>
        public List<Node> Subtitle { get; set; }
    }
}
