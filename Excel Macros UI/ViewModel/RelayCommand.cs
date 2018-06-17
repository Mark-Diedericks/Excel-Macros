/*
 * Mark Diedericks
 * 17/06/2015
 * Version 1.0.0
 * Relay command for actions
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace Excel_Macros_UI.ViewModel
{
    public class RelayCommand : ICommand
    {
        private readonly Action<object> ExecuteAction;
        private readonly Predicate<object> CanExecuteAction;

        public RelayCommand(Action<object> action) : this(action, _ => true)
        {

        }

        public RelayCommand(Action<object> action, Predicate<object> canExecute)
        {
            ExecuteAction = action;
            CanExecuteAction = canExecute;
        }

        public event EventHandler CanExecuteChanged
        {
            add
            {
                CommandManager.RequerySuggested += value;
            }

            remove
            {
                CommandManager.RequerySuggested -= value;
            }
        }

        public bool CanExecute(object parameter)
        {
            return CanExecuteAction(parameter);
        }

        public void Execute(object parameter)
        {
            ExecuteAction(parameter);
        }
    }
}
