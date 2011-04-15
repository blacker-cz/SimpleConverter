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
        /// Overlay specification
        /// </summary>
        private string _overlaySpec;

        /// <summary>
        /// Expanded overlay specification
        /// </summary>
        private ISet<int> _overlayList;

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
        public string OverlaySpec
        {
            set
            {
                _overlaySpec = value;

                // invalidate overlay list
                _overlayList = null;
            }
            get { return _overlaySpec; }
        }

        /// <summary>
        /// Overlay specification getter
        /// </summary>
        public ISet<int> OverlayList
        {
            get
            {
                if (_overlayList == null)
                {
                    _overlayList = Misc.ParseOverlay(_overlaySpec);
                }

                return _overlayList;
            }
        }

        /// <summary>
        /// Optional parameters
        /// </summary>
        public string OptionalParams { get; set; }

        /// <summary>
        /// Node content.
        /// Used only for leaves (type of content depends on type of node).
        /// </summary>
        public object Content { get; set; }

        /// <summary>
        /// List of children.
        /// </summary>
        public List<Node> Children { get; set; }

        #region Search implementation

        /// <summary>
        /// Next node id (used for search)
        /// </summary>
        private int _nextNode = 0;

        /// <summary>
        /// Node counter (used for search)
        /// </summary>
        private int _nodeCounter = 0;

        /// <summary>
        /// Last used search path in FindFirstNode
        /// </summary>
        private string _lastPath;

        /// <summary>
        /// Find first node by its path
        /// </summary>
        /// <param name="path">Node path, levels are separated by /</param>
        /// <example>var node = FindFirstNode("body/slide/string");</example>
        /// <returns>Node if found; null otherwise</returns>
        public Node FindFirstNode(string path)
        {
            _nextNode = 0;
            _lastPath = path;
            return FindNextNode(path);
        }

        /// <summary>
        /// Find next node
        /// </summary>
        /// <returns>Node if found; null otherwise</returns>
        public Node FindNextNode()
        {
            if (_lastPath == null)
                throw new InvalidOperationException("FindFirstNode not called!");

            return FindNextNode(_lastPath);
        }

        /// <summary>
        /// Find next node by its path (recursive)
        /// </summary>
        /// <param name="path">Node path</param>
        /// <returns>Node if found; null otherwise</returns>
        private Node FindNextNode(string path)
        {
            string[] nodesPath = path.Split(new char[] {'/'}, 2);

            foreach (var node in Children)
            {
                if (node.Type == path)
                {
                    if (nodesPath.Length != 1)
                    {
                        Node tmp =  FindNextNode(nodesPath[1]);
                        if (tmp != null)
                            return tmp;
                    }
                    else
                    {
                        if (_nextNode == _nodeCounter)
                        {
                            _nextNode++;
                            _nodeCounter = 0;

                            return node;
                        }
                        _nodeCounter++;
                    }
                }
            }

            _nodeCounter = 0;
            return null;
        }

        #endregion // Search implementation
    }
}
