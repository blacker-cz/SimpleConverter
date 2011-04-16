using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    /// <summary>
    /// Class containing settings from preambule of document
    /// </summary>
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
        /// Input file path
        /// </summary>
        public string InputDir { get; set; }

        /// <summary>
        /// Public constructor
        /// </summary>
        public PreambuleSettings(string inputDir)
        {
            TitlepageSettings = new Dictionary<string, List<Node>>();
            GraphicsPath = new HashSet<string>();
            InputDir = inputDir;
            GraphicsPath.Add(InputDir);
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
                    case "institute":
                        TitlepageSettings[node.Type] = node.Children;
                        break;
                    case "graphicspath":
                        foreach (Node child in node.Children)
                        {
                            if (child.Type == "path" && (child.Content as string).Length > 0)
                            {
                                string path = Path.Combine(InputDir, child.Content as string);

                                // add path only if directory exists
                                if(Directory.Exists(path))
                                    GraphicsPath.Add(child.Content as string);
                            }
                        }
                        break;
                    case "usepackage":
                        if (node.Content as string == "inputenc" && node.OptionalParams != "utf8")
                            Messenger.Instance.SendMessage("Unsupported code page, some characters may be broken", Contract.MessageLevel.WARNING);
                        break;
                    // unknown or invalid node -> ignore
                    default:
                        break;
                }
            }

        }
    }
}
