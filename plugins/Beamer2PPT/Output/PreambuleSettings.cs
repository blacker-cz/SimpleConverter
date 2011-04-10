using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    class PreambuleSettings
    {
        /// <summary>
        /// Table with global title page settings (author, title, date, etc.)
        /// </summary>
        public Dictionary<string, List<Node>> TitlepageSettings { get; set; }

        /// <summary>
        /// Set of paths where to look for images
        /// </summary>
        public ISet<string> GraphicsPath { get; set; }

        /// <summary>
        /// Public constructor
        /// </summary>
        public PreambuleSettings()
        {
            TitlepageSettings = new Dictionary<string, List<Node>>();
            GraphicsPath = new HashSet<string>();
        }

        /// <summary>
        /// Parse preabule and fill class properties by parsed values
        /// </summary>
        /// <param name="preambule">Preambule node</param>
        public void Parse(Node preambule)
        {
            if (preambule.Type != "preambule")
                throw new ArgumentException("Function accepts only preambule node type.");

            // process preambule nodes
            foreach (Node node in preambule.Children)
            {
                switch (node.Type)
                {
                    // title page settings
                    case "author":
                    case "title":
                    case "date":
                        TitlepageSettings[node.Type] = node.Children;
                        break;
                    case "graphicspath":
                        foreach (Node child in node.Children)
                        {
                            if (child.Type == "path" && (child.Content as string).Length > 0)
                                GraphicsPath.Add(child.Content as string);
                        }
                        break;
                    // unknown or invalid node -> ignore
                    default:
                        break;
                }
            }

        }
    }
}
