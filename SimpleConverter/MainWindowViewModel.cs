﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows.Input;
using System.Windows;

namespace SimpleConverter
{
    /// <summary>
    /// ViewModel for MainWindow.
    /// 
    /// Implements Model-View-ViewModel pattern.
    /// </summary>
    public class MainWindowViewModel : Contract.BaseViewModel
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
                         _browseCommand,
                         _stopBatchCommand,
                         _aboutCommand;

        /// <summary>
        /// Background thread object reference
        /// </summary>
        private BackgroundThread _backgroundThread;

        /// <summary>
        /// Public constructor.
        /// </summary>
        public MainWindowViewModel()
        {
            if (Factory.Loader.Instance.Plugins.Count<Contract.IPluginMetaData>() == 0)
                throw new Factory.PluginLoaderException("No plugins available.");

            Plugins = new ObservableCollection<Contract.IPluginMetaData>(Factory.Loader.Instance.Plugins);
            Messages = new Contract.AsyncObservableCollection<ListMessage>();
            Files = new Contract.ElementCollection<ListFile>();

            SelectedPlugin = Factory.Loader.Instance.Plugins.First<Contract.IPluginMetaData>();
            SelectPluginEnabled = true;
            SettingsTabEnabled = true;

            // load output path from user settings
            _outputPath = Properties.Settings.Default.OutputPath;

            // initialize background thread
            _backgroundThread = new BackgroundThread();
            _backgroundThread.ThreadEndedEvent += new ThreadEndedDelegate(ConversionDone);
            _backgroundThread.ConversionProgressEvent += new Contract.ProgressDelegate(OnCompleteProgress);
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
        public Contract.ElementCollection<ListFile> Files { get; private set; }

        #region Binding for button commands

        /// <summary>
        /// Command binding for start conversion command
        /// </summary>
        public BaseCommand StartConversionCommand { get { return _startConversionCommand ?? (_startConversionCommand = new StartConversionCommand(this)); } }

        /// <summary>
        /// Command binding for add file command
        /// </summary>
        public BaseCommand AddFileCommand { get { return _addFileCommand ?? (_addFileCommand = new AddFileCommand(this)); } }

        /// <summary>
        /// Command binding for remove file command
        /// </summary>
        public BaseCommand RemoveFileCommand { get { return _removeFileCommand ?? (_removeFileCommand = new RemoveFileCommand(this)); } }

        /// <summary>
        /// Command binding for browse command
        /// </summary>
        public BaseCommand BrowseCommand { get { return _browseCommand ?? (_browseCommand = new BrowseCommand(this)); } }

        /// <summary>
        /// Command binding for stop batch command
        /// </summary>
        public BaseCommand StopBatchCommand { get { return _stopBatchCommand ?? (_stopBatchCommand = new StopBatchCommand(this, true)); } }

        /// <summary>
        /// Command binding for about command
        /// </summary>
        public BaseCommand AboutCommand { get { return _aboutCommand ?? (_aboutCommand = new AboutCommand(this)); } }

        #endregion // Binding for button commands

        /// <summary>
        /// Allow drop flag
        /// </summary>
        public bool AllowDrop { get { return !AddFileCommand.Disabled; } }

        /// <summary>
        /// Output path textbox flag
        /// </summary>
        public bool OutputPathEnabled { get { return !BrowseCommand.Disabled; } }

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
        /// Flag for enabling/disabling select plugin combobox
        /// </summary>
        public bool SelectPluginEnabled { get; private set; }

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

        /// <summary>
        /// Progress counter for current file
        /// </summary>
        public int FileProgress { get; private set; }

        /// <summary>
        /// Progress counter for whole batch
        /// </summary>
        public int CompleteProgress { get; private set; }

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

        /// <summary>
        /// Flag for enabling/disabling settings tab
        /// </summary>
        public bool SettingsTabEnabled { get; private set; }

        /// <summary>
        /// Selected tab
        /// </summary>
        public int SelectedTab { get; set; }

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

        /// <summary>
        /// Event handling for complete progress change
        /// </summary>
        /// <param name="progress">Current progress</param>
        private void OnCompleteProgress(int progress)
        {
            CompleteProgress = progress;
            InvokePropertyChanged("CompleteProgress");
        }

        #endregion // Event handlers

        #region Button Click Handlers

        /// <summary>
        /// Convert button clicked
        /// </summary>
        public void StartConversionClicked()
        {
            Messages.Clear();   // clear message list

            try
            {
                if(!_backgroundThread.Run(_currentPlugin, Files, OutputPath))
                    MessageBox.Show("Couldn't start new thread.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch (Factory.InvalidArgumentException ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            AddFileCommand.Disabled = true;
            RemoveFileCommand.Disabled = true;
            StartConversionCommand.Disabled = true;
            BrowseCommand.Disabled = true;
            StopBatchCommand.Disabled = false;

            SelectPluginEnabled = false;
            SettingsTabEnabled = false;

            SelectedTab = 1;

            InvokePropertyChanged(null);
        }

        /// <summary>
        /// Conversion in background thread ended
        /// </summary>
        public void ConversionDone(int successful, int from)
        {
            AddFileCommand.Disabled = false;
            RemoveFileCommand.Disabled = false;
            StartConversionCommand.Disabled = false;
            BrowseCommand.Disabled = false;
            StopBatchCommand.Disabled = true;

            SelectPluginEnabled = true;
            SettingsTabEnabled = true;

            _backgroundThread.Join();

            InvokePropertyChanged(null);

            // show information message box
            string message = String.Format("Successfully converted {0} file(s) from {1}.\nFor more information look at progress messages.", successful, from);
            MessageBox.Show(message, "Conversion completed", MessageBoxButton.OK, successful == from ? MessageBoxImage.Information : MessageBoxImage.Warning);
        }

        /// <summary>
        /// Clicked on Add file button
        /// </summary>
        public void AddFileClicked()
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Multiselect = true;
            dlg.CheckFileExists = true;

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
        /// Add file to list (used by drop event)
        /// </summary>
        /// <param name="files">File path</param>
        public void AddFiles(string[] files)
        {
            if (AllowDrop)  // check if add file is allowed
            {
                foreach (string file in files)
                {
                    Files.Add(new ListFile(file, _currentPlugin.ValidateFile(file)));
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
            dlg.Description = "Select output directory.";
            dlg.ShowNewFolderButton = true;

            System.Windows.Forms.DialogResult result = dlg.ShowDialog();

            if(result == System.Windows.Forms.DialogResult.OK)
            {
                OutputPath = dlg.SelectedPath;
                InvokePropertyChanged("OutputPath");
            }

            dlg.Dispose();
        }

        /// <summary>
        /// About command handler
        /// </summary>
        public void About()
        {
            AboutBox about = new AboutBox(Application.Current.MainWindow);
            about.Title = @"About SimpleConverter";
            try
            {
                System.IO.StreamReader reader = new System.IO.StreamReader(Application.GetResourceStream(new System.Uri("/about.txt", UriKind.Relative)).Stream);
                about.description.Text = reader.ReadToEnd();
            }
            catch (System.IO.IOException)
            {
            }
            about.ShowDialog();
        }

        /// <summary>
        /// Stop background thread processing
        /// </summary>
        public void StopBatch()
        {
            _backgroundThread.Abort();
        }

        #endregion // Button Click Handlers
    }
}
