/*
 * Mark Diedericks
 * 22/07/2018
 * Version 1.0.6
 * The main hub of the interop library's connection to the UI
 */

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
        public delegate void ClearIOEvent();
        public event ClearIOEvent ClearAllIOEvent;

        public delegate void MacroAddEvent(Guid id, string macroName, string macroPath, Action macroClickEvent);
        public event MacroAddEvent AddRibbonMacroEvent;

        public delegate void MacroRemoveEvent(Guid id);
        public event MacroRemoveEvent RemoveRibbonMacroEvent;

        public delegate void MacroEditEvent(Guid id, string macroName, string macroPath);
        public event MacroEditEvent RenameRibbonMacroEvent;

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

        public static void AddRibbonMacro(Guid id, string macroName, string macroPath, Action macroClickEvent)
        {
            GetInstance().AddRibbonMacroEvent?.Invoke(id, macroName, macroPath, macroClickEvent);
        }

        public static void RemoveRibbonMacro(Guid id)
        {
            GetInstance().RemoveRibbonMacroEvent?.Invoke(id);
        }

        public static void RenameRibbonMacro(Guid id, string macroName, string macroPath)
        {
            GetInstance().RenameRibbonMacroEvent?.Invoke(id, macroName, macroPath);
        }

        public static void ClearAllIO()
        {
            GetInstance().ClearAllIOEvent?.Invoke();
        }

        #endregion
    }
}
