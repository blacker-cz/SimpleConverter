using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel.Composition;

namespace SimpleConverter.Contract
{
    /// <summary>
    /// MetaData definition class
    /// </summary>
    [MetadataAttribute]
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public class PluginMetadataAttribute : ExportAttribute
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="name">Plugin name</param>
        /// <param name="version">Plugin version</param>
        /// <param name="description">Plugin description</param>
        public PluginMetadataAttribute(string name, string version, string description = "")
            : base(typeof(IPluginMetaData))
        {
            Name = name;
            Version = version;
            Description = description;

            // automatically build key from plugin name and version
            Key = Hash.ComputeHash(name + version);
        }

        /// <summary>
        /// Plugin name
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Plugin version
        /// </summary>
        public string Version { get; private set; }

        /// <summary>
        /// Plugin description
        /// </summary>
        public string Description { get; private set; }

        /// <summary>
        /// Plugin key (generated automatically)
        /// </summary>
        public string Key { get; private set; }
    }
}
