using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace SimpleConverter.Plugin.Beamer2PPT
{
    /// <summary>
    /// ViewModel for SettingsView
    /// 
    /// Implements Model-View-ViewModel pattern.
    /// </summary>
    class SettingsViewViewModel : Contract.BaseViewModel
    {
        /// <summary>
        /// Constructor
        /// </summary>
        public SettingsViewViewModel()
        {
            SaveTypes = new ObservableCollection<SaveFileType>();

            #region Add save file types to collection

            SaveTypes.Add(new SaveFileType("Save in the default format.", PowerPoint.PpSaveAsFileType.ppSaveAsDefault));
            SaveTypes.Add(new SaveFileType("Save as a presentation (.pptx).", PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation));
            SaveTypes.Add(new SaveFileType("Save as a presentation 97-2003 (.ppt)", PowerPoint.PpSaveAsFileType.ppSaveAsPresentation));
            SaveTypes.Add(new SaveFileType("Save as a PDF.", PowerPoint.PpSaveAsFileType.ppSaveAsPDF));
            SaveTypes.Add(new SaveFileType("Save as an HTML document.", PowerPoint.PpSaveAsFileType.ppSaveAsHTML));
            SaveTypes.Add(new SaveFileType("Save as an BMP image.", PowerPoint.PpSaveAsFileType.ppSaveAsBMP));
            SaveTypes.Add(new SaveFileType("Save as a GIF image.", PowerPoint.PpSaveAsFileType.ppSaveAsGIF));
            SaveTypes.Add(new SaveFileType("Save as a JPG image.", PowerPoint.PpSaveAsFileType.ppSaveAsJPG));
            SaveTypes.Add(new SaveFileType("Save as a PNG image.", PowerPoint.PpSaveAsFileType.ppSaveAsPNG));

            #endregion // Add save file types to collection
        }

        /// <summary>
        /// Collection with possible save types
        /// </summary>
        public ObservableCollection<SaveFileType> SaveTypes { get; private set; }

        /// <summary>
        /// Selected save type
        /// </summary>
        public PowerPoint.PpSaveAsFileType SelectedSaveType
        {
            get { return Settings.Instance.SaveAs; }
            set
            {
                Settings.Instance.SaveAs = value;
            }
        }
    }

    /// <summary>
    /// Class wrapper for save file type info
    /// </summary>
    class SaveFileType
    {
        /// <summary>
        /// Name
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Type
        /// </summary>
        public PowerPoint.PpSaveAsFileType Type { get; private set; }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="name">Name</param>
        /// <param name="type">Type</param>
        public SaveFileType(string name, PowerPoint.PpSaveAsFileType type)
        {
            Name = name;
            Type = type;
        }
    }
}
