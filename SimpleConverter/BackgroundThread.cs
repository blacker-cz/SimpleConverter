using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using SimpleConverter.Contract;
using SimpleConverter.Factory;

namespace SimpleConverter
{
    /// <summary>
    /// Background batch conversion
    /// todo: maybe switch from threads to BackgroundWorker class
    /// </summary>
    class BackgroundThread
    {
        /// <summary>
        /// Private lock
        /// </summary>
        private object _lock = new Object();

        /// <summary>
        /// Abort flag
        /// </summary>
        private bool _abort;

        /// <summary>
        /// Plugin instance
        /// </summary>
        private IPlugin _plugin;

        /// <summary>
        /// List of files to convert
        /// </summary>
        private List<ListFile> _files;

        /// <summary>
        /// Output directory path
        /// </summary>
        private string _outputPath;

        /// <summary>
        /// Background thread
        /// </summary>
        private Thread _thread;

        /// <summary>
        /// Thread ended event
        /// </summary>
        public event ThreadEndedDelegate ThreadEndedEvent;

        /// <summary>
        /// Public constructor
        /// </summary>
        public BackgroundThread()
        {
        }

        /// <summary>
        /// Run background conversion
        /// </summary>
        /// <param name="plugin"></param>
        /// <param name="files"></param>
        /// <param name="outputPath"></param>
        /// <returns>true at success; false otherwise</returns>
        public bool Run(IPlugin plugin, IEnumerable<ListFile> files, string outputPath = "")
        {
            if (_thread != null && _thread.IsAlive)
                throw new InvalidStateException();

            else if (_thread != null && !_thread.IsAlive)
                _thread.Join();

            if (plugin == null)
                throw new InvalidArgumentException();

            if (files == null || files.Count<ListFile>() == 0)
                throw new InvalidArgumentException("Can't start conversion when there is no file to convert.");

            _plugin = plugin;

            // shallow copy of collection (deep would be better)
            _files = new List<ListFile>(files);

            _outputPath = outputPath;

            lock (_lock)
            {
                _abort = false;
            }

            _thread = new Thread(new ThreadStart(Worker));

            try
            {
                _thread.Start();
            }
            catch (ThreadStateException)
            {
                return false;
            }

            return true;
        }

        /// <summary>
        /// Background thread worker
        /// </summary>
        private void Worker()
        {
            foreach (ListFile file in _files)
            {
                if (file.Valid)
                {
#if !DEBUG
                    try
                    {
#endif
                        _plugin.ConvertDocument(file.Filepath, _outputPath);
#if !DEBUG
                    }
                    catch (Exception ex)
                    {
                        // todo: print error message
                        System.Windows.MessageBox.Show(ex.Message);
                        break;
                    }
#endif
                }

                lock (_lock)
                {
                    if (_abort)
                        break;
                }
            }

            // fire thread ended event
            try
            {
                ThreadEndedEvent();
            }
            catch (NullReferenceException)
            {
            }
        }

        /// <summary>
        /// Set abort conversion flag
        /// </summary>
        public void Abort()
        {
            lock (_lock)
            {
                _abort = true;
            }
        }
    }

    /// <summary>
    /// Delegate for thread ended event
    /// </summary>
    public delegate void ThreadEndedDelegate();
}
