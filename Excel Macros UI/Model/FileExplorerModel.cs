using Excel_Macros_UI.Model.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_UI.Model
{
    public class FileExplorerModel : ToolModel
    {
        private static FileExplorerModel s_Instance;

        public static FileExplorerModel GetInstance()
        {
            return s_Instance != null ? s_Instance : new FileExplorerModel();
        }

        public FileExplorerModel()
        {
            s_Instance = this;
        }
    }
}
