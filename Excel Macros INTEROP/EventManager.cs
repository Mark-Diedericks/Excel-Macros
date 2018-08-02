/*
 * Mark Diedericks
 * 02/08/2018
 * Version 1.0.8
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

        /// <summary>
        /// Gets instance of EventManager
        /// </summary>
        /// <returns>EventManager instance</returns>
        public static EventManager GetInstance()
        {
            return s_Instance;
        }

        /// <summary>
        /// Private instantiation of EventManager
        /// </summary>
        private EventManager()
        {
            s_Instance = this;
        }

        /// <summary>
        /// Public instantiation of EventManager
        /// </summary>
        public static void Instantiate()
        {
            new EventManager();
        }

        #endregion

        #region Event Firing

        /// <summary>
        /// Fires SetExcelIneractive event
        /// </summary>
        /// <param name="enabled">Whether or not Excel should be set as interactive</param>
        public static void ExcelSetInteractive(bool enabled)
        {
            GetInstance().SetExcelInteractive?.Invoke(enabled);
        }

        /// <summary>
        /// Fires FastWorkbook event
        /// </summary>
        /// <param name="enabled">Whether or not the ActiveWorkbook should enable or disable FastWorkbook</param>
        public static void FastWorkbook(bool enabled)
        {
            GetInstance().FastWorkbookEvent?.Invoke(enabled);
        }

        /// <summary>
        /// Fires FastWorksheet event
        /// </summary>
        /// <param name="worksheet">The worksheet which it should be applied to</param>
        /// <param name="enabled">Whether or not the worksheet should enable or disable FastWorksheet</param>
        public static void FastWorksheet(Excel.Worksheet worksheet, bool enabled)
        {
            GetInstance().FastWorksheetEvent?.Invoke(worksheet, enabled);
        }

        /// <summary>
        /// Fires AddRibbonMacro event
        /// </summary>
        /// <param name="id">The macro's id</param>
        /// <param name="macroName">The macro's name</param>
        /// <param name="macroPath">The macro's relative path</param>
        /// <param name="macroClickEvent">Event callback for when the macro is clicked</param>
        public static void AddRibbonMacro(Guid id, string macroName, string macroPath, Action macroClickEvent)
        {
            GetInstance().AddRibbonMacroEvent?.Invoke(id, macroName, macroPath, macroClickEvent);
        }

        /// <summary>
        /// Fires RemoveRibbonMacro event
        /// </summary>
        /// <param name="id">The macro's id</param>
        public static void RemoveRibbonMacro(Guid id)
        {
            GetInstance().RemoveRibbonMacroEvent?.Invoke(id);
        }

        /// <summary>
        /// Fires RenameRibbonMacro event
        /// </summary>
        /// <param name="id">The macro's id</param>
        /// <param name="macroName">The macro's name</param>
        /// <param name="macroPath">The macro's relative path</param>
        public static void RenameRibbonMacro(Guid id, string macroName, string macroPath)
        {
            GetInstance().RenameRibbonMacroEvent?.Invoke(id, macroName, macroPath);
        }

        /// <summary>
        /// Fires ClearAllIO event
        /// </summary>
        public static void ClearAllIO()
        {
            GetInstance().ClearAllIOEvent?.Invoke();
        }

        #endregion
    }
}
