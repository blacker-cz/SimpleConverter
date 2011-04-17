using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace SimpleConverter
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            // attach ViewModel
            try
            {
                this.DataContext = new MainWindowViewModel();
            }
            catch (Factory.PluginLoaderException e)
            {
                this.Close();
                MessageBox.Show("Application encountered following error and will now end:\n\n\"" + e.Message + "\"", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Drop event handler
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void listBoxFiles_Drop(object sender, DragEventArgs e)
        {
            MainWindowViewModel mv = this.DataContext as MainWindowViewModel;

            if( e.Data.GetDataPresent(DataFormats.FileDrop, false) == true )
                mv.AddFiles((string[]) e.Data.GetData(DataFormats.FileDrop));
        }
    }
}
