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
        /// <summary>
        /// ViewModel associated with command
        /// </summary>
        protected readonly MainWindowViewModel _viewModel;

        /// <summary>
        /// Flag if control is disabled or not
        /// </summary>
        private bool _disabled;

        /// <summary>
        /// Getter/setter for <see cref="_disabled"/> flag
        /// </summary>
        public bool Disabled
        {
            get { return _disabled; }
            set
            {
                if (_disabled != value)
                {
                    _disabled = value;
                    CommandManager.InvalidateRequerySuggested();
                }
            }
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="viewModel">Associated ViewModel</param>
        /// <param name="disabled">Flag if control is disabled</param>
        public BaseCommand(MainWindowViewModel viewModel, bool disabled = false)
        {
            if (viewModel == null)
                throw new ArgumentNullException();

            _viewModel = viewModel;
            Disabled = disabled;
        }

        /// <summary>
        /// Execute method for command
        /// </summary>
        /// <param name="parameter">Parameter</param>
        public abstract void Execute(object parameter);

        /// <summary>
        /// CanExecute method for command
        /// </summary>
        /// <param name="parameter">Parameter</param>
        /// <returns>Return value depends on <see cref="_disabled"/> flag</returns>
        public bool CanExecute(object parameter)
        {
            return !Disabled;
        }

        /// <summary>
        /// CanExecuteChanged event
        /// </summary>
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
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="viewModel">Associated ViewModel</param>
        /// <param name="disabled">Flag if control is disabled</param>
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
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="viewModel">Associated ViewModel</param>
        /// <param name="disabled">Flag if control is disabled</param>
        public AddFileCommand(MainWindowViewModel viewModel, bool disabled = false) : base(viewModel, disabled) { }

        /// <summary>
        /// Execute method for command
        /// </summary>
        /// <param name="parameter">Parameter</param>
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
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="viewModel">Associated ViewModel</param>
        /// <param name="disabled">Flag if control is disabled</param>
        public RemoveFileCommand(MainWindowViewModel viewModel, bool disabled = false) : base(viewModel, disabled) { }

        /// <summary>
        /// Execute method for command
        /// </summary>
        /// <param name="parameter">Parameter</param>
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
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="viewModel">Associated ViewModel</param>
        /// <param name="disabled">Flag if control is disabled</param>
        public BrowseCommand(MainWindowViewModel viewModel, bool disabled = false) : base(viewModel, disabled) { }

        /// <summary>
        /// Execute method for command
        /// </summary>
        /// <param name="parameter">Parameter</param>
        public override void Execute(object parameter)
        {
            _viewModel.BrowseClicked();
        }
    }

    /// <summary>
    /// Command handler class for Stop batch button
    /// </summary>
    public class StopBatchCommand : BaseCommand
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="viewModel">Associated ViewModel</param>
        /// <param name="disabled">Flag if control is disabled</param>
        public StopBatchCommand(MainWindowViewModel viewModel, bool disabled = false) : base(viewModel, disabled) { }

        /// <summary>
        /// Execute method for command
        /// </summary>
        /// <param name="parameter">Parameter</param>
        public override void Execute(object parameter)
        {
            _viewModel.StopBatch();
        }
    }

    /// <summary>
    /// Command handler class for About button
    /// </summary>
    public class AboutCommand : BaseCommand
    {
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="viewModel">Associated ViewModel</param>
        /// <param name="disabled">Flag if control is disabled</param>
        public AboutCommand(MainWindowViewModel viewModel, bool disabled = false) : base(viewModel, disabled) { }

        /// <summary>
        /// Execute method for command
        /// </summary>
        /// <param name="parameter">Parameter</param>
        public override void Execute(object parameter)
        {
            _viewModel.About();
        }
    }
}
