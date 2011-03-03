using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Input;

namespace SimpleConverter
{
    public class StartConversionCommand : ICommand
    {
        private readonly MainWindowViewModel _viewModel;

        public StartConversionCommand(MainWindowViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        public void Execute(object parameter)
        {
            _viewModel.StartConversionClicked();
        }

        public bool CanExecute(object parameter)
        {
            // todo: implement this
            return true;
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
    }
}
