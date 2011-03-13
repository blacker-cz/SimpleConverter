using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace SimpleConverter.Contract
{

    /// <summary>
    /// Observable collection with method for raising CollectionChanged event
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ElementCollection<T> : ObservableCollection<T>
    {
        /// <summary>
        /// Public method for raising CollectionChanged event
        /// </summary>
        public void UpdateCollection()
        {
            OnCollectionChanged(new System.Collections.Specialized.NotifyCollectionChangedEventArgs(
                                System.Collections.Specialized.NotifyCollectionChangedAction.Reset));
        }
    }
}
