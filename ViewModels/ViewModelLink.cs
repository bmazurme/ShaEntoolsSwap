using System;
using System.ComponentModel;
using System.Globalization;
using System.Threading;
using System.Windows.Input;

namespace Entools.ViewLinkModels
{
    class ViewModelLink : INotifyPropertyChanged
    {
        #region FIELDS
        private string _linkUrl;
        public string LinkUrl
        {
            get
            {
                return _linkUrl;
            }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {                  
                    Properties.Settings.Default["LinkUrl"] = _linkUrl = value;
                    OnPropertyChanged("LinkUrl");
                    string sheetId = value.Replace("https://docs.google.com/spreadsheets/d/", "");
                    sheetId = sheetId.Substring(0, sheetId.IndexOf("/"));
                    Properties.Settings.Default["Sheet"] = sheetId;
                    Properties.Settings.Default.Save();
                }
            }
        }
        #endregion

        public ViewModelLink()
        {
            #region INTERFACE

            CultureInfo ui = Thread.CurrentThread.CurrentUICulture;

            if (ui.Name == "ru-RU") System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("ru-RU");
            else System.Threading.Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US");

            #endregion

            _linkUrl = Properties.Settings.Default["LinkUrl"].ToString();
            OnPropertyChanged("LinkUrl");
        }

        public bool CanExecute
        {
            get
            {
                // check if executing is allowed, i.e., validate, check if a process is running, etc. 
                //return true / false;
                return true;
            }
        }

        public Action CloseAction { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }


    public class CommandHandler : ICommand
    {
        private Action _action;
        private Func<bool> _canExecute;

        /// <summary>
        /// Creates instance of the command handler
        /// </summary>
        /// <param name="action">Action to be executed by the command</param>
        /// <param name="canExecute">A bolean property to containing current permissions to execute the command</param>
        public CommandHandler(Action action, Func<bool> canExecute)
        {
            _action = action;
            _canExecute = canExecute;
        }


        /// <summary>
        /// Wires CanExecuteChanged event 
        /// </summary>
        public event EventHandler CanExecuteChanged
        {
            add { CommandManager.RequerySuggested += value; }
            remove { CommandManager.RequerySuggested -= value; }
        }


        /// <summary>
        /// Forcess checking if execute is allowed
        /// </summary>
        /// <param name="parameter"></param>
        /// <returns></returns>
        public bool CanExecute(object parameter)
        {
            return _canExecute.Invoke();
        }


        public void Execute(object parameter)
        {
            _action();
        }
    }
}