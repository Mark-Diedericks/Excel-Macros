/*
 * Mark Diedericks
 * 17/06/2015
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
    public class FileExplorerViewModel : ToolViewModel
    {
        public FileExplorerViewModel()
        {
            Model = new FileExplorerModel();
        }

        #region Model

        public new FileExplorerModel Model
        {
            get
            {
                return (FileExplorerModel)base.Model;
            }

            set
            {
                if (((FileExplorerModel)base.Model) != value)
                {
                    base.Model = value;
                    OnPropertyChanged(nameof(Model));
                }
            }
        }

        #endregion

    }
}
