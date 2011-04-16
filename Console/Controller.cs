using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SimpleConverter.Console
{
    class Controller
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public Controller()
        {
            if (Factory.Loader.Instance.Plugins.Count<Contract.IPluginMetaData>() == 0)
                throw new Factory.PluginLoaderException("No plugins available.");
        }

        /// <summary>
        /// List available plugins
        /// </summary>
        public void ListPlugins()
        {
            System.Console.WriteLine("Listing available plugins:\n");
            System.Console.WriteLine("{0,10}   {1,-45}{2,10}", "Key", "Plugin name", "Version");
            System.Console.WriteLine("-----------------------------------------------------------------------");

            foreach (Contract.IPluginMetaData plugin in Factory.Loader.Instance.Plugins)
            {
                System.Console.WriteLine("{0,10}   {1,-45}{2,10}", plugin.Key, plugin.Name, plugin.Version);
            }
        }

        /// <summary>
        /// Convert files
        /// </summary>
        /// <param name="plugin_key">Plugin key</param>
        /// <param name="output_dir">Output directory</param>
        /// <param name="extra">List of files to convert</param>
        /// <returns>0 if successfull; otherwise error number</returns>
        public int Convert(string plugin_key, string output_dir, List<string> extra)
        {
            throw new NotImplementedException();
        }
    }
}
