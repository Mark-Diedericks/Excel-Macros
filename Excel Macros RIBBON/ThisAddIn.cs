﻿/*
 * Mark Diedericks
 * 02/08/2018
 * Version 1.0.4
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
        private UI.Routing.EventManager m_EventManager;

        private bool m_RibbonLoaded = false;
        private bool m_ApplicationLoaded = false;
        private Thread m_Thread;

        private delegate void CloseEvent();
        private static event CloseEvent ApplicationClosing;

        /// <summary>
        /// Event Callback for when the AddIn is started by Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            if (ExcelMacrosRibbonTab.GetInstance() == null)
                ExcelMacrosRibbonTab.MacroRibbonLoadEvent += MacroRibbonLoaded;
            else
                MacroRibbonLoaded();

            UI.Routing.EventManager.ApplicationLoadedEvent += MacroEditorLoaded;

            Dispatcher dispatcher = Dispatcher.CurrentDispatcher;

            m_Thread = new Thread(() =>
            {
                UI.Routing.EventManager.CreateApplicationInstance(Application, dispatcher, Properties.Settings.Default.RibbonMacros);
            });

            m_Thread.SetApartmentState(ApartmentState.STA);
            m_Thread.IsBackground = true;
            m_Thread.Start();
        }

        /// <summary>
        /// Event Callback for when the AddIn is shutdown by Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Properties.Settings.Default.RibbonMacros = m_RibbonTab.GetRibbonMacros();
            Properties.Settings.Default.Save();

            if (m_EventManager != null)
                m_EventManager.Shutdown();

            ApplicationClosing?.Invoke();
            
            try
            {
                m_Thread.Join();
            }
            catch(Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Event Callback for when the Ribbon is fully loaded, sets events
        /// </summary>
        private void MacroRibbonLoaded()
        {
            m_RibbonLoaded = true;
            m_RibbonTab = ExcelMacrosRibbonTab.GetInstance();

            if (m_ApplicationLoaded)
                SetEvents();
        }

        /// <summary>
        /// Event Callback for when the Editor is fully loaded, sets events
        /// </summary>
        private void MacroEditorLoaded()
        {
            m_ApplicationLoaded = true;
            m_EventManager = UI.Routing.EventManager.GetInstance();

            if (m_RibbonLoaded)
                SetEvents();
        }

        /// <summary>
        /// Binds events between the AddIn, the RibbonTab as well as the UI's EventManager
        /// </summary>
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

            m_EventManager.AddRibbonMacroEvent += m_RibbonTab.AddMacro;
            m_EventManager.RemoveRibbonMacroEvent += m_RibbonTab.RemoveMacro;
            m_EventManager.RenameRibbonMacroEvent += m_RibbonTab.RenameMacro;

            m_EventManager.SetExcelInteractiveEvent += (enable) =>
            {
                if (Application.Interactive == enable)
                    return;

                Application.Interactive = enable;
            };

            //FastWorksheet Macro by Paul Bica
            //https://stackoverflow.com/questions/30959315/excel-vba-performance-1-million-rows-delete-rows-containing-a-value-in-less
            m_EventManager.FastWorkbookEvent += (enable) =>
            {
                Application.Calculation = enable ? Excel.XlCalculation.xlCalculationManual : Excel.XlCalculation.xlCalculationAutomatic;
                Application.DisplayAlerts = !enable;
                Application.DisplayStatusBar = !enable;
                Application.EnableAnimations = !enable;
                Application.EnableEvents = !enable;
                Application.ScreenUpdating = !enable;

                foreach (Excel.Worksheet ws in Application.ActiveWorkbook.Sheets)
                {
                    ws.DisplayPageBreaks = false;
                    ws.EnableCalculation = !enable;
                    ws.EnableFormatConditionsCalculation = !enable;
                    ws.EnablePivotTable = !enable;
                }
            };

            //FastWorksheet Macro by Paul Bica
            //https://stackoverflow.com/questions/30959315/excel-vba-performance-1-million-rows-delete-rows-containing-a-value-in-less
            m_EventManager.FastWorksheetEvent += (worksheet, enable) =>
            {
                worksheet.DisplayPageBreaks = false;
                worksheet.EnableCalculation = !enable;
                worksheet.EnableFormatConditionsCalculation = !enable;
                worksheet.EnablePivotTable = !enable;
            };

            m_EventManager.DisplayInputMessageEvent += (message, title, def, left, top, helpFile, helpContextID, type, OnResult) =>
            {
                object result = Application.InputBox(message, title, def, left, top, helpFile, helpContextID, type);
                OnResult?.Invoke(result);
            };

            m_EventManager.DisplayInputMessageReturnEvent += (message, title, def, left, top, helpFile, helpContextID, type) =>
            {
                return Application.InputBox(message, title, def, left, top, helpFile, helpContextID, type);
            };

            m_RibbonTab.MainUILoaded();
            UI.Routing.EventManager.MacroRibbonLoaded();
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
