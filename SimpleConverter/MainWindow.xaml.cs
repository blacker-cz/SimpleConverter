﻿using System;
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
            catch (Exception e) // todo fix this :)
            {
                this.Close();
                MessageBox.Show("Application encountered following error:\n\n\"" + e.Message + "\"", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
