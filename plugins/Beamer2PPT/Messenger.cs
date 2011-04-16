using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SimpleConverter.Contract;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    /// <summary>
    /// Composite class for messengers implementation.
    /// 
    /// Implements composite and singleton patterns.
    /// todo: think about moving messenger (with interface) to contract
    /// </summary>
    sealed class Messenger : IMessenger
    {
        /// <summary>
        /// Singleton instance
        /// </summary>
        private static volatile Messenger instance;

        /// <summary>
        /// Lock object
        /// </summary>
        private static object syncRoot = new Object();

        /// <summary>
        /// List of registered messengers
        /// </summary>
        private List<IMessenger> _messengers;

        /// <summary>
        /// Private constructor
        /// </summary>
        private Messenger()
        {
            _messengers = new List<IMessenger>();
        }

        /// <summary>
        /// Public instance property
        /// </summary>
        public static Messenger Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (syncRoot)
                    {
                        if(instance == null)
                            instance = new Messenger();
                    }
                }

                return instance;
            }
        }

        #region Implementation of IMessenger

        /// <summary>
        /// Send message
        /// </summary>
        /// <param name="message">Message text</param>
        /// <param name="level">Message level</param>
        public void SendMessage(string message, MessageLevel level = MessageLevel.INFO)
        {
            foreach (IMessenger item in _messengers)
            {
                try
                {
                    item.SendMessage(message, level);
                }
                catch (NullReferenceException)
                {
                }
            }
        }

        #endregion

        /// <summary>
        /// Register messenger
        /// </summary>
        /// <param name="messenger">Messenger instance</param>
        public void Add(IMessenger messenger)
        {
            _messengers.Add(messenger);
        }

        /// <summary>
        /// Unregister messenger
        /// </summary>
        /// <param name="messenger">Messenger instance</param>
        /// <returns>true if item is successfully removed; otherwise, false</returns>
        public bool Remove(IMessenger messenger)
        {
            return _messengers.Remove(messenger);
        }
    }
}
