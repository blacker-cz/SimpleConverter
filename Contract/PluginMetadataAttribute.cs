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
        public PluginMetadataAttribute(string name, string version, string description = "")
            : base(typeof(IPluginMetaData))
        {
            Name = name;
            Version = version;
            Description = description;

            // automatically build key from plugin name and version
            Key = Hash.md5(name + version);
        }

        public string Name { get; private set; }
        public string Version { get; private set; }
        public string Description { get; private set; }
        public string Key { get; private set; }
    }
}
