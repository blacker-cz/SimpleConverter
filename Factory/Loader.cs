using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel.Composition;
using System.ComponentModel.Composition.Hosting;
using SimpleConverter.Contract;

namespace SimpleConverter.Factory
{
    /// <summary>
    /// Class for loading plugins.
    /// 
    /// Implements singleton pattern.
    /// </summary>
    public class Loader
    {
        /// <summary>
        /// Dynamically loaded plugins
        /// </summary>
        [ImportMany]
        private Lazy<IPlugin, IPluginMetaData>[] LoadedPlugins { get; set; }

        /// <summary>
        /// Singleton instance.
        /// </summary>
        private static Loader instance;

        /// <summary>
        /// Private constructor.
        /// </summary>
        private Loader()
        {
            try
            {
                var catalog = new AggregateCatalog();
                catalog.Catalogs.Add(new DirectoryCatalog(@".\plugins"));
                var container = new CompositionContainer(catalog);
                container.ComposeParts(this);
            }
            catch
            {
                throw new PluginLoaderException("Couldn't load plugins.");
            }
        }

        /// <summary>
        /// Public instance property.
        /// </summary>
        public static Loader Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new Loader();
                }

                return instance;
            }
        }

        /// <summary>
        /// List of plugins metadata.
        /// </summary>
        public IEnumerable<IPluginMetaData> Plugins
        {
            get
            {
                foreach (var plugin in LoadedPlugins)
                {
                    yield return plugin.Metadata;
                }
            }
        }

        /// <summary>
        /// Public getter for plugin instance.
        /// </summary>
        /// <param name="index">Plugin unique key</param>
        /// <returns>Instance of IPlugin if found, null otherwise</returns>
        public IPlugin this[string index]
        {
            get 
            {
                foreach (var plugin in LoadedPlugins)
                {
                    if (plugin.Metadata.Key == index)
                        return plugin.Value;
                }

                return null;
            }
        }
    }
}
