using System;
using System.Windows.Input;

namespace AccessCodeLib.AccUnit.VbeAddIn
{
    public interface IButtonCommand : ICommand
    {
        string Caption { get; }    
    }

    public class ButtonRelayCommand : RelayCommand, IButtonCommand
    {
        public ButtonRelayCommand(Action execute, string caption, Func<bool> canExecute = null) 
            : base(execute, canExecute)
        {
            Caption = caption;  
        }

        public string Caption { get; } 
    }

    public class RelayCommand : ICommand
    {
        private readonly Action execute;
        private readonly Func<bool> canExecute;

        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }

        public RelayCommand(Action execute, Func<bool> canExecute = null)
        {
            this.execute = execute ?? throw new ArgumentNullException(nameof(execute));
            this.canExecute = canExecute;
        }

        public bool CanExecute(object parameter)
        {
            return canExecute == null || canExecute();
        }

        public void Execute(object parameter)
        {
            execute();
        }
    }
}
