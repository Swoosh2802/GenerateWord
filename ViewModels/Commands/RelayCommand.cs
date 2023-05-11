using System;
using System.Windows.Input;

namespace Bindings.Commands
{
    public class ParameterizedRelayCommand<T> : ICommand
    {
        public static ParameterizedRelayCommand<T> From(Action<T> action, Predicate<T> canExecute) => new ParameterizedRelayCommand<T>()
        {
            Action = action,
            Verify = canExecute
        };

        public static ParameterizedRelayCommand<T> From(Action<T> action) => From(action, T => true);

        private Predicate<T> Verify { get; set; }

        private Action<T> Action { get; set; }

        public bool CanExecute(object parameter) => Verify((T)parameter);

        public void Execute(object parameter) => Action((T)parameter);

        public event EventHandler CanExecuteChanged;

        public void RaiseExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, new EventArgs());
        }
    }

    public class ParameterlessRelayCommand : ICommand
    {
        public static ParameterlessRelayCommand From(Action action, Func<bool> canExecute) => new ParameterlessRelayCommand()
        {
            Action = action,
            Verify = canExecute
        };

        public static ParameterlessRelayCommand From(Action action) => From(action, () => true);

        private Func<bool> Verify { get; set; }

        private Action Action { get; set; }

        public bool CanExecute(object parameter) => Verify();

        public void Execute(object parameter) => Action();

        public event EventHandler CanExecuteChanged;

        public void RaiseExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, new EventArgs());
        }
    }
}