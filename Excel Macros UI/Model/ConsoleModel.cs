using Excel_Macros_UI.Model.Base;
using Excel_Macros_UI.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace Excel_Macros_UI.Model
{
    public class ConsoleModel : ToolModel
    {
        private static ConsoleModel s_Instance;

        public static ConsoleModel GetInstance()
        {
            return s_Instance != null ? s_Instance : new ConsoleModel();
        }

        public ConsoleModel()
        {
            s_Instance = this;
            Output = new TextBoxWriter(null);
            Error = new TextBoxWriter(null);

            PreferredLocation = PaneLocation.Bottom;
        }

        private TextBoxWriter m_Output;
        public TextBoxWriter Output
        {
            get
            {
                return m_Output;
            }
            set
            {
                m_Output = value;
                OnPropertyChanged(nameof(Output));
            }
        }

        private TextBoxWriter m_Error;
        public TextBoxWriter Error
        {
            get
            {
                return m_Error;
            }
            set
            {
                m_Error = value;
                OnPropertyChanged(nameof(Error));
            }
        }
    }
}
