using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    /// <summary>
    /// Plugin settings class.
    /// 
    /// Implements singleton pattern.
    /// </summary>
    sealed class Settings
    {
        #region Singleton implementation

        /// <summary>
        /// Singleton instance.
        /// </summary>
        private static volatile Settings instance;

        /// <summary>
        /// Synchronization node used for lock
        /// </summary>
        private static object _syncRoot = new Object();

        /// <summary>
        /// Public instance property.
        /// </summary>
        public static Settings Instance
        {
            get
            {
                if (instance == null)
                {
                    lock (_syncRoot)
                    {
                        if (instance == null)
                            instance = new Settings();
                    }
                }

                return instance;
            }
        }

        #endregion // Singleton implementation

        /// <summary>
        /// Private constructor.
        /// </summary>
        private Settings()
        {
            SaveAs = PowerPoint.PpSaveAsFileType.ppSaveAsDefault;
            AdjustSize = true;
            NestedAsText = true;
        }

        /// <summary>
        /// File saving format
        /// </summary>
        public PowerPoint.PpSaveAsFileType SaveAs { get; set; }

        /// <summary>
        /// Adjust units length from beamer to PowerPoint (PowerPoint has double slide size than beamer)
        /// </summary>
        public bool AdjustSize { get; set; }

        /// <summary>
        /// Type of processing nested elements.
        /// True - from nested elements process only formatted text, overlays, pause commands
        /// False - generate nested elements at the end of slide, ignore pause commands
        /// </summary>
        public bool NestedAsText { get; set; }
    }
}
