using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Input;

namespace SimpleConverter
{
    /// <summary>
    /// Base class for commands implementation
    /// </summary>
    public abstract class BaseCommand : ICommand
    {
        protected readonly MainWindowViewModel _viewModel;

        public bool Disabled { get; set; }

        public BaseCommand(MainWindowViewModel viewModel, bool disabled = false)
        {
            _viewModel = viewModel;
            Disabled = disabled;
        }

        public abstract void Execute(object parameter);

        public bool CanExecute(object parameter)
        {
            return !Disabled;
        }

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }
    }

    /// <summary>
    /// Command handler class for Start conversion button
    /// </summary>
    public class StartConversionCommand : BaseCommand
    {
        public StartConversionCommand(MainWindowViewModel viewModel, bool disabled = false) : base(viewModel, disabled) { }

        public override void Execute(object parameter)
        {
            _viewModel.StartConversionClicked();
        }
    }

    /// <summary>
    /// Command handler class for Add file button
    /// </summary>
    public class AddFileCommand : BaseCommand
    {
        public AddFileCommand(MainWindowViewModel viewModel, bool disabled = false) : base(viewModel, disabled) { }

        public override void Execute(object parameter)
        {
            _viewModel.AddFileClicked();
        }
    }

    /// <summary>
    /// Command handler class for Remove file button
    /// </summary>
    public class RemoveFileCommand : BaseCommand
    {
        public RemoveFileCommand(MainWindowViewModel viewModel, bool disabled = false) : base(viewModel, disabled) { }

        public override void Execute(object parameter)
        {
            _viewModel.RemoveFileClicked();
        }
    }

    /// <summary>
    /// Command handler class for Browse button
    /// </summary>
    public class BrowseCommand : BaseCommand
    {
        public BrowseCommand(MainWindowViewModel viewModel, bool disabled = false) : base(viewModel, disabled) { }

        public override void Execute(object parameter)
        {
            _viewModel.BrowseClicked();
        }
    }
}
