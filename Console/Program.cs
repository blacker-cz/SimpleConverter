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
                System.Console.Write("SimpleConverter: ");
                System.Console.WriteLine("No parameters given.");
                System.Console.WriteLine("Try `sc --help' for more information.");
                return 1;
            }

            var p = new OptionSet() {
			    { "p|plugin=", "{KEY} of used conversion plugin.",  v => plugin_key = v },
			    { "o|output=", "output {DIRECTORY} for converted files.", v => output_dir = v },
			    { "h|help",  "show this message and exit", v => show_help = v != null },
			    { "l|list",  "list available plugins and exit", v => list_plugins = v != null },
            };

            List<string> extra;
            try
            {
                extra = p.Parse(args);
            }
            catch (OptionException e)
            {
                System.Console.Write("SimpleConverter: ");
                System.Console.WriteLine(e.Message);
                System.Console.WriteLine("Try `sc --help' for more information.");
                return 1;
            }

            // print help
            if (show_help)
            {
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
                System.Console.Write("SimpleConverter: ");
                System.Console.WriteLine("no plugin key given.");
                System.Console.WriteLine("Try `sc --help' for more information.");
                return 1;
            }

            return controller.Convert(plugin_key, output_dir, extra);
        }

    }
}
