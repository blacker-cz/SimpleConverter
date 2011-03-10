using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    /// <summary>
    /// Class for document Node
    /// </summary>
    public class Node
    {

#if DEBUG
        /// <summary>
        /// Parameterless constructor for serialization/deserialization (debugging purposes only)
        /// </summary>
        public Node() { }
#endif

        /// <summary>
        /// Constructor implementation
        /// </summary>
        public Node(string type, string overlay = "", string optional = "", List<Node> children = null, object content = null)
        {
            Type = type;
            OverlaySpec = overlay;
            OptionalParams = optional;
            Children = children;
            Content = content;
        }

        /// <summary>
        /// Node type (e.g. bold, slide).
        /// Content depends only on input filter.
        /// </summary>
        public string Type { get; set; }

        /// <summary>
        /// Overlay specification
        /// </summary>
        public string OverlaySpec { get; set; }

        /// <summary>
        /// Optional parameters
        /// </summary>
        public string OptionalParams { get; set; }

        /// <summary>
        /// Node content.
        /// Used only for leaves (type of content depends on type of node).
        /// todo: change to string?
        /// </summary>
        public object Content { get; set; }

        /// <summary>
        /// List of children.
        /// </summary>
        public List<Node> Children { get; set; }
    }
}
