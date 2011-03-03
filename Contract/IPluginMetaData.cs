using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SimpleConverter.Contract
{
    /// <summary>
    /// Public interface for MetaData classes
    /// </summary>
    public interface IPluginMetaData
    {
        /// <summary>
        /// Plugin name
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Plugin version
        /// </summary>
        string Version { get; }

        /// <summary>
        /// Plugin description
        /// </summary>
        string Description { get; }

        /// <summary>
        /// Plugin unique key
        /// </summary>
        string Key { get; }
    }
}
