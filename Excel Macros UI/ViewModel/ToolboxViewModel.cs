using Excel_Macros_UI.Model;
using Excel_Macros_UI.ViewModel.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_UI.ViewModel
{
    public class ToolboxViewModel : ToolViewModel
    {
        public ToolboxViewModel()
        {
            Model = new ToolboxModel();
        }

        #region Model

        public new ToolboxModel Model
        {
            get
            {
                return (ToolboxModel)base.Model;
            }

            set
            {
                if (((ToolboxModel)base.Model) != value)
                {
                    base.Model = value;
                    OnPropertyChanged(nameof(Model));
                }
            }
        }

        #endregion

    }
}
