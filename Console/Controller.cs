using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace SimpleConverter.Console
{
    /// <summary>
    /// Simple controller for console interface
    /// </summary>
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
        /// <returns>0 if successful; otherwise error number</returns>
        public int Convert(string plugin_key, string output_dir, List<string> extra)
        {
            int returnCode = 0;

            Contract.IPlugin plugin = Factory.Loader.Instance[plugin_key];

            if (plugin == null)     // check if plugin exists
            {
                Program.PrintError("Plugin with key '{0}' not found.", plugin_key);
                return 1;
            }

            // if output directory is not set, use default (and print info)
            if (output_dir == null || output_dir.Length == 0)
            {
                output_dir = Path.Combine(Directory.GetCurrentDirectory(), "output");
                try
                {
                    if (!Directory.Exists(output_dir))
                        Directory.CreateDirectory(output_dir);
                }
                catch (Exception)
                {
                    Program.PrintError("Couldn't create output directory");
                    return 1;
                }
                System.Console.WriteLine("Output directory not set. Using '{0}'", output_dir);
            }

            // register message handler
            plugin.SendMessageEvent += new Contract.SendMessageDelegate(plugin_SendMessageEvent);

            try
            {
                plugin.Init();

                foreach (string file in extra)
                {
                    try
                    {
                        if (plugin.ValidateFile(file))
                        {
                            System.Console.WriteLine("-----------------------------------------------------------------------");
                            System.Console.WriteLine("Converting file '{0}'", file);

                            plugin.ConvertDocument(file, output_dir);
                        }
                    }
                    catch (Contract.DocumentException) { }
                }

            }
            // methods raising these exceptions should add message to log, so no need to do anything in here
            catch (Contract.InitException) { returnCode = 1; }
            finally
            {
                plugin.Done();
            }

            System.Console.WriteLine("-----------------------------------------------------------------------");

            return returnCode;
        }

        /// <summary>
        /// Set plugin options
        /// </summary>
        /// <param name="plugin_key">Plugin key</param>
        /// <param name="options">List of options</param>
        /// <returns>Extra</returns>
        public List<string> SetPluginOptions(string plugin_key, List<string> options)
        {
            Contract.IPlugin plugin = Factory.Loader.Instance[plugin_key];

            if (plugin == null)     // check if plugin exists
            {
                Program.PrintError("Plugin with key '{0}' not found.", plugin_key);
                return null;
            }

            return plugin.ConsoleOptions(options);
        }

        /// <summary>
        /// Print plugin help
        /// </summary>
        /// <param name="plugin_key"></param>
        /// <returns>true if plugin exists; false otherwise</returns>
        public bool PrintPluginHelp(string plugin_key)
        {
            Contract.IPlugin plugin = Factory.Loader.Instance[plugin_key];

            if (plugin == null)     // check if plugin exists
            {
                Program.PrintError("Plugin with key '{0}' not found.", plugin_key);
                return false;
            }

            plugin.ConsoleHelp();

            return true;
        }

        /// <summary>
        /// Print message to console
        /// </summary>
        /// <param name="message">Message text</param>
        /// <param name="level">Level of message</param>
        void plugin_SendMessageEvent(string message, Contract.MessageLevel level = Contract.MessageLevel.INFO)
        {
            System.Console.WriteLine("{0,-10}{1}", level.ToString() + ":", message);
        }
    }
}
