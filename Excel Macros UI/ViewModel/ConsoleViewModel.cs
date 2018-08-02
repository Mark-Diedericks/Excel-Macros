/*
 * Mark Diedericks
 * 22/07/2018
 * Version 1.0.1
 * Console view model
 */
 
using Excel_Macros_UI.Model;
using Excel_Macros_UI.Utilities;
using Excel_Macros_UI.ViewModel.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_UI.ViewModel
{
    public class ConsoleViewModel : ToolViewModel
    {
        /// <summary>
        /// Instantiate ConsoleViewModel
        /// </summary>
        public ConsoleViewModel()
        {
            Model = new ConsoleModel();
        }

        #region Model

        public new ConsoleModel Model
        {
            get
            {
                return (ConsoleModel)base.Model;
            }

            set
            {
                if (((ConsoleModel)base.Model) != value)
                {
                    base.Model = value;
                    OnPropertyChanged(nameof(Model));
                }
            }
        }
        
        public TextBoxWriter Output
        {
            get
            {
                return Model.Output;
            }
            set
            {
                Model.Output = value;
                OnPropertyChanged(nameof(Output));
            }
        }
        
        public TextBoxWriter Error
        {
            get
            {
                return Model.Error;
            }
            set
            {
                Model.Error = value;
                OnPropertyChanged(nameof(Error));
            }
        }

        #endregion

    }
}
