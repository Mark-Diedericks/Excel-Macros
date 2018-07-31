using Excel_Macros_UI.Model.Base;
using Excel_Macros_UI.Utilities;
using Excel_Macros_UI.View;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_UI.Model
{
    public class ToolboxModel : ToolModel
    {
        private static ToolboxModel s_Instance;

        public static ToolboxModel GetInstance()
        {
            return s_Instance != null ? s_Instance : new ToolboxModel();
        }

        public ToolboxModel()
        {
            s_Instance = this;
            Routing.EventManager.DocumentChangedEvent += DocumentChangedEvent;
        }

        private void DocumentChangedEvent(ViewModel.Base.DocumentViewModel vm)
        {
            //Change between textual and visual toolbox items
            //throw new NotImplementedException();
        }
    }
}
