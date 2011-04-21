using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NDesk.Options;

namespace SimpleConverter.Console
{
    class Program
    {
        static int Main(string[] args)
        {
            bool show_help = false, list_plugins = false;
            string plugin_key = "", output_dir = "";

            // no parameters given -> end execution
            if (args.Length == 0)
            {
                PrintError("No parameters given.");
                return 1;
            }

            var p = new OptionSet() {
			    { "p|plugin=", "{KEY} of used conversion plugin.",  v => plugin_key = v },
			    { "o|output=", "output {DIRECTORY} for converted files.", v => output_dir = v },
			    { "h|?|help",  "show this message and exit (if plugin is specified show plugin help)", v => show_help = v != null },
			    { "l|list",  "list available plugins and exit", v => list_plugins = v != null },
            };

            List<string> extra;
            try
            {
                extra = p.Parse(args);
            }
            catch (OptionException e)
            {
                PrintError(e.Message);
                return 1;
            }

            // print help (if plugin key not specified)
            if (show_help && (plugin_key == null || plugin_key == ""))
            {
                System.Console.WriteLine("SimpleConverter - universal document converter");
                System.Console.WriteLine("Copyright (c) 2011 Lukáš Černý");
                System.Console.WriteLine();
                System.Console.WriteLine("Usage: sc [OPTIONS]+ files+");
                System.Console.WriteLine("Convert document files using specified plugin.");
                System.Console.WriteLine();
                System.Console.WriteLine("Options:");
                p.WriteOptionDescriptions(System.Console.Out);
                return 0;
            }

            Controller controller;

            try
            {
                controller = new Controller();
            }
            catch (Factory.PluginLoaderException ex)
            {
                System.Console.WriteLine(ex.Message);
                return 1;
            }

            // list plugins
            if (list_plugins)
            {
                controller.ListPlugins();
                return 0;
            }

            if (plugin_key == null || plugin_key.Length == 0)
            {
                PrintError("No plugin key given.");
                return 1;
            }

            // show plugin help
            if (show_help)
            {
                System.Console.WriteLine("SimpleConverter - universal document converter");
                System.Console.WriteLine("Copyright (c) 2011 Lukáš Černý");
                System.Console.WriteLine();
                System.Console.WriteLine("Plugin options:");
                if (!controller.PrintPluginHelp(plugin_key))
                    return 1;
                else
                    return 0;
            }

            // process plugin options
            extra = controller.SetPluginOptions(plugin_key, extra);
            if (extra == null)
                return 1;

            try
            {
                return controller.Convert(plugin_key, output_dir, extra);
            }
            catch (Exception ex)
            {
                System.Console.Error.WriteLine("Application encountered following unrecoverable error and will now exit:\n\n\"" + ex.Message + "\"");
                return 1;
            }
        }

        /// <summary>
        /// Print error message
        /// </summary>
        /// <param name="message">Message to print</param>
        /// <param name="arg">Message arguments</param>
        public static void PrintError(string message, params object[] arg)
        {
            System.Console.Error.Write("SimpleConverter: ");
            System.Console.Error.WriteLine(message, arg);
            System.Console.Error.WriteLine("Try `sc --help' for more information.");
        }

    }
}
