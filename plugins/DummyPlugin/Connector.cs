using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel.Composition;
using SimpleConverter.Contract;

namespace SimpleConverter.Plugin.Dummy
{
    [Export(typeof(IPlugin))]
    [PluginMetadata("Dummy Convertor", "1.0.0.1", "This convertor does nothing. And has very long description. Lorem ipsum etc.")]
    public class Connector : IPlugin
    {
        private System.Windows.FrameworkElement _visual;

        public event SendMessageDelegate SendMessageEvent;

        public event ProgressDelegate ProgressEvent;

        private Random _random;

        public Connector()
        {
            _random = new Random(DateTime.Now.Millisecond);
        }

        public void Init()
        {
        }

        public void Done()
        {
        }

        public void ConvertDocument(string filename, string outputDirectory = "")
        {
            throw new NotImplementedException();
        }

        public System.Windows.FrameworkElement GetVisual()
        {
            if (_visual == null)    // keep plugin view instantiated
                _visual = new SettingsView();

            return _visual;
        }

        public bool ValidateFile(string filename)
        {
            return (_random.Next() % 2 == 1);
        }
    }
}
