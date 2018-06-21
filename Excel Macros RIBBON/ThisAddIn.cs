/*
 * Mark Diedericks
 * 21/06/2015
 * Version 1.0.1
 * The main hub of the Excel AddIn -> used only for the ribbon tab it allows me to add
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using UI = Excel_Macros_UI;
using Microsoft.Office.Tools.Excel;
using System.Windows.Threading;
using System.Threading;

namespace Excel_Macros_RIBBON
{
    public partial class ThisAddIn
    {
        private ExcelMacrosRibbonTab m_RibbonTab;
        private UI.EventManager m_EventManager;

        private bool m_RibbonLoaded = false;
        private bool m_ApplicationLoaded = false;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            ExcelMacrosRibbonTab.MacroRibbonLoadEvent += MacroRibbonLoaded;
            UI.EventManager.ApplicationLoaded += MacroEditorLoaded;

            Dispatcher excelDispatcher = Dispatcher.CurrentDispatcher;
            UI.EventManager.CreateApplicationInstance(Application, excelDispatcher);
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            this.Application.Interactive = true;

            if (m_EventManager != null)
                m_EventManager.ShutdownEvent();
        }

        private void MacroRibbonLoaded()
        {
            m_RibbonLoaded = true;
            m_RibbonTab = ExcelMacrosRibbonTab.GetInstance();

            if (m_ApplicationLoaded)
                SetEvents();
        }

        private void MacroEditorLoaded()
        {
            m_ApplicationLoaded = true;
            m_EventManager = UI.EventManager.GetInstance();

            if (m_RibbonLoaded)
                SetEvents();
        }

        private void SetEvents()
        {
            if(m_RibbonTab == null || m_EventManager == null)
            {
                System.Diagnostics.Debug.WriteLine("RibbonTab: " + m_RibbonTab);
                System.Diagnostics.Debug.WriteLine("EventManager: " + m_EventManager);
                Environment.Exit(1);
            }

            m_RibbonTab.MacroEditorClickEvent += m_EventManager.MacroEditorClickEvent;
            m_RibbonTab.NewTextualClickEvent += m_EventManager.NewTextualClickEvent;
            m_RibbonTab.NewVisualClickEvent += m_EventManager.NewVisualClickEvent;
            m_RibbonTab.OpenMacroClickEvent += m_EventManager.OpenMacroClickEvent;

            m_EventManager.AddRibbonMacro += m_RibbonTab.AddMacro;
            m_EventManager.RemoveRibbonMacro += m_RibbonTab.RenameMacro;
            m_EventManager.RenameRibbonMacro += m_RibbonTab.RenameMacro;

            m_RibbonTab.MainUILoaded();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
