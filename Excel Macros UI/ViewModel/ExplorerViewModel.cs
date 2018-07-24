/*
 * Mark Diedericks
 * 17/06/2018
 * Version 1.0.0
 * File explorer view model
 */

using Excel_Macros_UI.Model;
using Excel_Macros_UI.ViewModel.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_UI.ViewModel
{
    public class ExplorerViewModel : ToolViewModel
    {
        public ExplorerViewModel()
        {
            Model = new ExplorerModel();
        }

        #region Model

        public new ExplorerModel Model
        {
            get
            {
                return (ExplorerModel)base.Model;
            }

            set
            {
                if (((ExplorerModel)base.Model) != value)
                {
                    base.Model = value;
                    OnPropertyChanged(nameof(Model));
                }
            }
        }

        #endregion

    }
}
