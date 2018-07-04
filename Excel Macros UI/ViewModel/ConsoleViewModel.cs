using Excel_Macros_UI.Model;
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

        #endregion

    }
}
