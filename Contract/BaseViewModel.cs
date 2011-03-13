using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace SimpleConverter.Contract
{
    /// <summary>
    /// Base ViewModel class
    /// </summary>
    public abstract class BaseViewModel : INotifyPropertyChanged
    {
        #region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        protected void InvokePropertyChanged(string propertyName)
        {
            var e = new PropertyChangedEventArgs(propertyName);

            PropertyChangedEventHandler changed = PropertyChanged;

            if (changed != null)
                changed(this, e);
        }

        #endregion  // Implementation of INotifyPropertyChanged
    }
}
