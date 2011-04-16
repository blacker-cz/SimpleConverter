﻿using System;
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
        /// Synchronization context
        /// </summary>
        private SynchronizationContext _synchronizationContext = SynchronizationContext.Current;

        /// <summary>
        /// Public constructor
        /// </summary>
        public BackgroundThread()
        {
        }

        /// <summary>
        /// Run background conversion
        /// </summary>
        /// <param name="plugin">Plugin instance</param>
        /// <param name="files">List of files to convert</param>
        /// <param name="outputPath">Output folder for generated files</param>
        /// <returns>true at success; false otherwise</returns>
        public bool Run(IPlugin plugin, IEnumerable<ListFile> files, string outputPath = "")
        {
            if (_thread != null && _thread.IsAlive)
                throw new InvalidStateException();

            else if (_thread != null && !_thread.IsAlive)
                _thread.Join();

            if (plugin == null)
                throw new InvalidArgumentException("There is a problem with plugin selection, please try reselect plugin.");

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

            try
            {
                _plugin.Init();

                foreach (ListFile file in _files)
                {
                    try
                    {
                        if (file.Valid)
                        {
                            _plugin.ConvertDocument(file.Filepath, _outputPath);
                        }
                    }
                    catch (DocumentException) { }

                    lock (_lock)
                    {
                        if (_abort)
                            break;
                    }
                }

            }
            // methods raising these exceptions should add message to log, so no need to do anything in here
            catch (InitException) { }
            finally
            {
                _plugin.Done();
            }

            // fire thread ended event
            try
            {
                if (SynchronizationContext.Current == _synchronizationContext)
                {
                    // Execute the ThreadEndedEvent event on the current thread
                    ThreadEndedEvent();
                }
                else
                {
                    // Post the ThreadEndedEvent event on the creator thread
                    _synchronizationContext.Post(new SendOrPostCallback(delegate(object state)
                    {
                        ThreadEndedDelegate handler = ThreadEndedEvent;

                        if (handler != null)
                        {
                            handler();
                        }
                    }), null);
                }
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

        /// <summary>
        /// Join thread
        /// </summary>
        public void Join()
        {
            if (_thread != null)
                _thread.Join();

            _thread = null;
        }
    }

    /// <summary>
    /// Delegate for thread ended event
    /// </summary>
    public delegate void ThreadEndedDelegate();
}
