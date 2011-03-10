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

        /// <summary>
        /// Output directory
        /// </summary>
        private string _outputPath;

        /// <summary>
        /// Command handlers
        /// </summary>
        private BaseCommand _startConversionCommand,
                         _addFileCommand,
                         _removeFileCommand,
                         _browseCommand;

        /// <summary>
        /// Public constructor.
        /// </summary>
        public MainWindowViewModel()
        {
            if (Factory.Loader.Instance.Plugins.Count<Contract.IPluginMetaData>() == 0)
                throw new Factory.PluginLoaderException("No plugins available.");

            Plugins = new ObservableCollection<Contract.IPluginMetaData>(Factory.Loader.Instance.Plugins);
            Messages = new ObservableCollection<ListMessage>();
            Files = new ElementCollection<ListFile>();
            SelectedPlugin = Factory.Loader.Instance.Plugins.First<Contract.IPluginMetaData>();

            _startConversionCommand = new StartConversionCommand(this);
            _addFileCommand = new AddFileCommand(this);
            _removeFileCommand = new RemoveFileCommand(this);
            _browseCommand = new BrowseCommand(this);

            // load output path from user settings
            _outputPath = Properties.Settings.Default.OutputPath;
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
        /// Collection of files
        /// </summary>
        public ElementCollection<ListFile> Files { get; private set; }

        /// <summary>
        /// Handlers for button click
        /// </summary>
        public ICommand StartConversionCommand { get { return _startConversionCommand; } }

        public ICommand AddFileCommand { get { return _addFileCommand; } }

        public ICommand RemoveFileCommand { get { return _removeFileCommand; } }

        public ICommand BrowseCommand { get { return _browseCommand;  } }

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

                // unregister events
                if (_currentPlugin != null)
                {
                    _currentPlugin.SendMessageEvent -= new Contract.SendMessageDelegate(OnSendMessage);
                    _currentPlugin.ProgressEvent -= new Contract.ProgressDelegate(OnFileProgress);
                }

                _currentPlugin = Factory.Loader.Instance[value.Key];

                // register events with currently selected plugin
                _currentPlugin.SendMessageEvent += new Contract.SendMessageDelegate(OnSendMessage);
                _currentPlugin.ProgressEvent += new Contract.ProgressDelegate(OnFileProgress);

                // revalidate files in list
                foreach (var item in Files)
                {
                    item.Valid = _currentPlugin.ValidateFile(item.Filepath);
                }
                Files.UpdateCollection();

                // inform view about plugin change
                InvokePropertyChanged("PluginView");
            }
        }

        /// <summary>
        /// Output directory path
        /// </summary>
        public string OutputPath
        {
            get { return _outputPath; }
            set
            {
                _outputPath = value;
                // remember path in user settings
                Properties.Settings.Default.OutputPath = _outputPath;
                Properties.Settings.Default.Save();
            }
        }

        public int FileProgress { get; private set; }

        /// <summary>
        /// Selected file in file list
        /// </summary>
        public ListFile SelectedFile { get; set; }

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

        #region Event handlers

        /// <summary>
        /// Event handling for send message event from plugin
        /// </summary>
        /// <param name="message">Message text</param>
        /// <param name="level">Message level</param>
        private void OnSendMessage(string message, Contract.MessageLevel level = Contract.MessageLevel.INFO)
        {
            Messages.Add(new ListMessage(message, level));
        }

        /// <summary>
        /// Event handling for file progress change
        /// </summary>
        /// <param name="progress">Current progress</param>
        private void OnFileProgress(int progress)
        {
            FileProgress = progress;
            InvokePropertyChanged("FileProgress");
        }

        #endregion // Event handlers

        #region Button Click Handlers

        public void StartConversionClicked()
        {
            // hardcoded for debugging purposes - process in new thread
            _currentPlugin.ConvertDocument(@"D:\Programovani\VS.2010\SimpleConverter\_bordel\beamer.tex");
        }

        /// <summary>
        /// Clicked on Add file button
        /// </summary>
        public void AddFileClicked()
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Multiselect = true;
            dlg.CheckFileExists = true;
            // todo: set file extensions depending on plugin?

            // Show open file dialog box
            bool? result = dlg.ShowDialog();

            // Process open file dialog box results
            if (result == true)
            {
                foreach (var item in dlg.FileNames)
                {
                    Files.Add(new ListFile(item, _currentPlugin.ValidateFile(item)));
                }
            }
        }

        /// <summary>
        /// Remove file from list button/key clicked
        /// </summary>
        public void RemoveFileClicked()
        {
            Files.Remove(SelectedFile);
        }

        /// <summary>
        /// Browse for output folder
        /// </summary>
        public void BrowseClicked()
        {
            // WPF doesn't have folder browser dialog, so we have to use the one from Windows.Forms
            System.Windows.Forms.FolderBrowserDialog dlg = new System.Windows.Forms.FolderBrowserDialog();
            dlg.SelectedPath = OutputPath;
            dlg.Description = "Select output directory.";   // todo: better description
            dlg.ShowNewFolderButton = true;

            System.Windows.Forms.DialogResult result = dlg.ShowDialog();

            if(result == System.Windows.Forms.DialogResult.OK)
            {
                OutputPath = dlg.SelectedPath;
                InvokePropertyChanged("OutputPath");
            }

            dlg.Dispose();
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
                        Icon = (System.Windows.Media.ImageSource)System.Windows.Application.Current.FindResource("iconInfo");
                        break;
                    case SimpleConverter.Contract.MessageLevel.WARNING:
                        Icon = (System.Windows.Media.ImageSource)System.Windows.Application.Current.FindResource("iconWarning");
                        break;
                    case SimpleConverter.Contract.MessageLevel.ERROR:
                        Icon = (System.Windows.Media.ImageSource)System.Windows.Application.Current.FindResource("iconError");
                        break;
                    default:
                        throw new ArgumentException("Unknown message level.");
                }
            }

            public string Message { get; private set; }

            public System.Windows.Media.ImageSource Icon { get; private set; }
        }

        #endregion // Message wrapper class

        #region File wrapper class

        public class ListFile
        {
            public ListFile(string filename, bool valid = false)
            {
                Filename = System.IO.Path.GetFileName(filename); ;
                Filepath = filename;
                Valid = valid;
            }

            public string Filename { get; private set; }

            public string Filepath { get; private set; }

            public bool Valid { get; set; }

            public string ValidColor
            {
                get
                {
                    if(Valid)
                        return "PaleGreen";
                    else
                        return "LightSalmon";
                }
            }
        }

        #endregion // File wrapper class
    }

    /// <summary>
    /// Observable collection with method for raising CollectionChanged event
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ElementCollection<T> : ObservableCollection<T>
    {
        public void UpdateCollection()
        {
            OnCollectionChanged(new System.Collections.Specialized.NotifyCollectionChangedEventArgs(
                                System.Collections.Specialized.NotifyCollectionChangedAction.Reset));
        }
    }
}
