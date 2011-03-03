using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows.Input;

namespace SimpleConverter
{
    /// <summary>
    /// ViewModel for MainWindow.
    /// 
    /// Implements Model-View-ViewModel pattern.
    /// </summary>
    public class MainWindowViewModel : INotifyPropertyChanged
    {
        /// <summary>
        /// Metadata of currently selected plugin.
        /// </summary>
        private Contract.IPluginMetaData _selectedPlugin;

        /// <summary>
        /// Object of currently selected plugin.
        /// </summary>
        private Contract.IPlugin _currentPlugin;

        private ICommand _startConversionCommand;

        /// <summary>
        /// Public constructor.
        /// </summary>
        public MainWindowViewModel()
        {
            if(Factory.Loader.Instance.Plugins.Count<Contract.IPluginMetaData>() == 0)
                throw new Exception("No plugins available."); // todo use better (appropriate) exception

            Plugins = new ObservableCollection<Contract.IPluginMetaData>(Factory.Loader.Instance.Plugins);
            Messages = new ObservableCollection<ListMessage>();
            SelectedPlugin = Factory.Loader.Instance.Plugins.First<Contract.IPluginMetaData>();

            _startConversionCommand = new StartConversionCommand(this);
        }

        /// <summary>
        /// Content of combobox with plugins
        /// </summary>
        public ObservableCollection<Contract.IPluginMetaData> Plugins { get; private set; }

        /// <summary>
        /// Collection of messages from plugin
        /// </summary>
        public ObservableCollection<ListMessage> Messages { get; private set; }

        /// <summary>
        /// Handler for button click
        /// </summary>
        public ICommand StartConversionCommand { get { return _startConversionCommand; } }

        /// <summary>
        /// Selected plugin in combobox
        /// </summary>
        public Contract.IPluginMetaData SelectedPlugin
        {
            get { return _selectedPlugin; }
            set
            {
                if (value == _selectedPlugin)
                    return;

                _selectedPlugin = value;

                // unregister message event
                if(_currentPlugin != null)
                    _currentPlugin.SendMessageEvent -= new Contract.SendMessageDelegate(OnSendMessage);

                _currentPlugin = Factory.Loader.Instance[value.Key];

                // register message even with currently selected plugin
                _currentPlugin.SendMessageEvent += new Contract.SendMessageDelegate(OnSendMessage);

                // inform view about plugin change
                InvokePropertyChanged("PluginView");
            }
        }

        /// <summary>
        /// Selected plugin settings view
        /// </summary>
        public System.Windows.FrameworkElement PluginView
        {
            get
            {
                return _currentPlugin.GetVisual();
            }
        }

        /// <summary>
        /// Event handling for send message event from plugin
        /// </summary>
        /// <param name="message">Message text</param>
        /// <param name="level">Message level</param>
        private void OnSendMessage(string message, Contract.MessageLevel level = Contract.MessageLevel.INFO)
        {
            Messages.Add(new ListMessage(message, level));
            // todo: invoke property change?
            //InvokePropertyChanged("Messages");
        }

        #region Button Click Handlers

        public void StartConversionClicked()
        {
            // hardcoded for debugging purposes
            _currentPlugin.ConvertDocument(@"D:\Programovani\VS.2010\SimpleConverter\_bordel\beamer.tex");
        }

        #endregion

        #region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        private void InvokePropertyChanged(string propertyName)
        {
            var e = new PropertyChangedEventArgs(propertyName);

            PropertyChangedEventHandler changed = PropertyChanged;

            if (changed != null)
                changed(this, e);
        }

        #endregion  // Implementation of INotifyPropertyChanged

        #region Message wrapper class

        public class ListMessage
        {
            public ListMessage(string message, Contract.MessageLevel level = Contract.MessageLevel.INFO)
            {
                Message = message;
                switch (level)
                {
                    case SimpleConverter.Contract.MessageLevel.INFO:
                        Icon = (System.Windows.Media.ImageSource) System.Windows.Application.Current.FindResource("iconInfo");
                        break;
                    case SimpleConverter.Contract.MessageLevel.WARNING:
                        Icon = (System.Windows.Media.ImageSource) System.Windows.Application.Current.FindResource("iconWarning");
                        break;
                    case SimpleConverter.Contract.MessageLevel.ERROR:
                        Icon = (System.Windows.Media.ImageSource) System.Windows.Application.Current.FindResource("iconError");
                        break;
                    default:
                        throw new ArgumentException("Unknown message level.");
                }
            }

            public string Message { get; private set; }

            public System.Windows.Media.ImageSource Icon { get; private set; }
        }

        #endregion
    }
}
