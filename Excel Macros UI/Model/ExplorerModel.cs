using Excel_Macros_UI.Model.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_UI.Model
{
    public class ExplorerModel : ToolModel
    {
        private static ExplorerModel s_Instance;

        public static ExplorerModel GetInstance()
        {
            return s_Instance != null ? s_Instance : new ExplorerModel();
        }

        public ExplorerModel()
        {
            s_Instance = this;
        }
    }
}
