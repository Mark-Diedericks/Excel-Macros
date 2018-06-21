using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Macros_INTEROP
{
    public class EventManager
    {

        //public delegate void MacroAddEvent(Guid id, string macroName, string macroPath, Action macroClickEvent);
        //public event MacroAddEvent AddRibbonMacro;

        public delegate void SetEnabled(bool enabled);
        public event SetEnabled SetExcelInteractive;
        public event SetEnabled FastWorkbookEvent;

        public delegate void SetEnabledWorksheet(Excel.Worksheet worksheet, bool enabled);
        public event SetEnabledWorksheet FastWorksheetEvent;

        #region Static Access & Instantiation

        private static EventManager s_Instance;

        public static EventManager GetInstance()
        {
            return s_Instance;
        }

        private EventManager()
        {
            s_Instance = this;
        }

        public static void Instantiate()
        {
            new EventManager();
        }

        #endregion

        #region Event Firing

        public static void ExcelSetInteractive(bool enabled)
        {
            GetInstance().SetExcelInteractive?.Invoke(enabled);
        }

        public static void FastWorkbook(bool enabled)
        {
            GetInstance().FastWorkbookEvent?.Invoke(enabled);
        }

        public static void FastWorksheet(Excel.Worksheet worksheet, bool enabled)
        {
            GetInstance().FastWorksheetEvent?.Invoke(worksheet, enabled);
        }

        #endregion
    }
}
