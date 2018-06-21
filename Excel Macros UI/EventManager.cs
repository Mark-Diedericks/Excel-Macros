/*
 * Mark Diedericks
 * 17/06/2015
 * Version 1.0.1
 * Event manager, allowing for cross-thread interaction between the Excel Ribbon tab and the UI/Interop projects
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Excel_Macros_INTEROP;
using Excel_Macros_INTEROP.Macros;
using System.Windows.Threading;
using System.Threading;

namespace Excel_Macros_UI
{
    public class EventManager
    {
        public delegate void MacroAddEvent(Guid id, string macroName, string macroPath, Action macroClickEvent);
        public event MacroAddEvent AddRibbonMacro;

        public delegate void MacroEditEvent(Guid id, string macroName, string macroPath);
        public event MacroEditEvent RemoveRibbonMacro;
        public event MacroEditEvent RenameRibbonMacro;

        public delegate void LoadEvent();
        public static event LoadEvent ApplicationLoaded;

        private static EventManager s_Instance;

        private EventManager()
        {
            s_Instance = this;
        }

        public static EventManager GetInstance()
        {
            return s_Instance;
        }

        public static void CreateApplicationInstance(Excel.Application application, Dispatcher uiDispatcher)
        {
            new EventManager();

            MainWindow.CreateInstance();

            Main.Initialize(application, uiDispatcher, new Action(() =>
            {
                Main.GetInstance().OnFocused += WindowFocusEvent;
                Main.GetInstance().OnShown += WindowShowEvent;
                Main.GetInstance().OnHidden += WindowHideEvent;

                MessageManager.GetInstance().DisplayOkMessageEvent += DisplayOkMessage;
                MessageManager.GetInstance().DisplayYesNoMessageEvent += DisplayYesNoMessage;

                LoadCompleted();
            }));
        }

        #region Excel to UI Events

        public void ShutdownEvent()
        {
            if (MainWindow.GetInstance() != null)
            {
                MainWindow.GetInstance().SaveAll();
                MainWindow.GetInstance().Hide();
            }

            Main.Destroy();
        }

        public void MacroEditorClickEvent()
        {
            if (MainWindow.GetInstance() == null)
                return;
            
            MainWindow.GetInstance().ShowWindow();
        }

        public void NewTextualClickEvent()
        {
            if (MainWindow.GetInstance() == null)
                return;

            MainWindow.GetInstance().ShowWindow();
            //MainWindow.GetInstance().CreateMacroAsync(null, MacroType.PYTHON, "/");
        }

        public void NewVisualClickEvent()
        {
            if (MainWindow.GetInstance() == null)
                return;

            MainWindow.GetInstance().ShowWindow();
            //MainWindow.GetInstance().CreateMacroAsync(null, MacroType.BLOCKLY, "/");
        }

        public void OpenMacroClickEvent()
        {
            if (MainWindow.GetInstance() == null)
                return;

            MainWindow.GetInstance().ShowWindow();
            //MainWindow.GetInstance().ImportMacroAsync(null, "/");
        }

        #endregion

        #region Main to Ribbon Events

        public void AddMacro(Guid id, IMacro macro)
        {
            MacroDeclaration md = Main.GetDeclaration(id);
            AddRibbonMacro?.Invoke(id, md.name, md.relativepath, delegate () { macro.ExecuteRelease(null, false); });
        }

        public void RemoveMacro(Guid id)
        {
            MacroDeclaration md = Main.GetDeclaration(id);
            RemoveRibbonMacro?.Invoke(id, md.name, md.relativepath);
        }

        public void RenameMacro(Guid id, IMacro macro)
        {
            MacroDeclaration md = Main.GetDeclaration(id);
            RenameRibbonMacro?.Invoke(id, md.name, md.relativepath);
        }

        public static void LoadCompleted()
        {
            ApplicationLoaded?.Invoke();
        }

        #endregion

        #region Main to UI Events

        private static void WindowFocusEvent()
        {
            MainWindow.GetInstance().TryFocus();
        }

        private static void WindowShowEvent()
        {
            MainWindow.GetInstance().ShowWindow();
        }

        private static void WindowHideEvent()
        {
            MainWindow.GetInstance().HideWindow();
        }

        private static void DisplayOkMessage(string content, string title)
        {
            MainWindow.GetInstance().DisplayOkMessage(content, title);
        }

        private static void DisplayYesNoMessage(string content, string title, Action<bool> OnReturn)
        {
            MainWindow.GetInstance().DisplayYesNoMessage(content, title, OnReturn);
        }

        #endregion
    }
}
