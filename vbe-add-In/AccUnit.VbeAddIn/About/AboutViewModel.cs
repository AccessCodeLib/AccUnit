using AccessCodeLib.AccUnit.Configuration;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Input;

namespace AccessCodeLib.AccUnit.VbeAddIn.About
{
    public class AboutViewModel
    {
        public ICommand NavigateCommand { get; }

        public string AddInVersion => AddInManager.FileVersion;
        public string FrameworkVersion => Configurator.FileVersion;

        public AboutViewModel()
        {
            NavigateCommand = new RelayCommand<string>(Navigate);
            Contributors = new List<Contributor>
            {
                new Contributor("Josef Pötzl"),
                new Contributor("Paul Rohorzka"),
                new Contributor("Sten Schmidt")
            };      
        }

        private void Navigate(string url)
        {
            if (Uri.TryCreate(url, UriKind.Absolute, out var uri))
            {
                Process.Start(new ProcessStartInfo(uri.AbsoluteUri) { UseShellExecute = true });
            }
        }

        public IEnumerable<Contributor> Contributors { get; }

    }

    public class RelayCommand<T> : ICommand
    {
        private readonly Action<T> _execute;
        private readonly Predicate<T> _canExecute;

        public RelayCommand(Action<T> execute, Predicate<T> canExecute = null)
        {
            _execute = execute ?? throw new ArgumentNullException(nameof(execute));
            _canExecute = canExecute;
        }

        public bool CanExecute(object parameter)
        {
            return _canExecute == null || _canExecute((T)parameter);
        }

        public void Execute(object parameter)
        {
            _execute((T)parameter);
        }

        public event EventHandler CanExecuteChanged
        {
            add => CommandManager.RequerySuggested += value;
            remove => CommandManager.RequerySuggested -= value;
        }
    }

    public class Contributor
    {
        public Contributor(string name)
        {
            Name = name;
        }

        public string Name { get; }
    }

}
