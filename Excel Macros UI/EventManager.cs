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

        public static void CreateApplicationInstance(Excel.Application application, Dispatcher interopDispatcher, Dispatcher excelDispatcher)
        {
            new EventManager();

            Thread WindowThread = new Thread((ThreadStart)delegate
            {
                Dispatcher windowDispatcher = Dispatcher.CurrentDispatcher;
                MainWindow.CreateInstance();
                
                interopDispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() => Main.Initialize(application, windowDispatcher, interopDispatcher, excelDispatcher, new Action(() =>
                {
                    Main.GetInstance().OnFocused += WindowFocusEvent;
                    Main.GetInstance().OnShown += WindowShowEvent;
                    Main.GetInstance().OnHidden += WindowHideEvent;

                    MessageManager.GetInstance().DisplayOkMessageEvent += DisplayOkMessage;
                    MessageManager.GetInstance().DisplayYesNoMessageEvent += DisplayYesNoMessage;

                    interopDispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() => EventManager.LoadCompleted()));
                }))));

                Dispatcher.Run();
            });

            WindowThread.SetApartmentState(ApartmentState.STA);
            WindowThread.IsBackground = true;
            WindowThread.Priority = ThreadPriority.Normal;
            WindowThread.Start();
        }

        #region Excel to UI Events

        public void ShutdownEvent()
        {
            if (MainWindow.GetInstance() != null)
                Main.GetWindowDispatcher().BeginInvoke(new Action(() => MainWindow.GetInstance().Close()));

            if (Main.GetWindowDispatcher() != null)
                Main.GetWindowDispatcher().BeginInvokeShutdown(DispatcherPriority.Send);

            Main.Destroy();
        }

        public void MacroEditorClickEvent()
        {
            if (MainWindow.GetInstance() == null)
                return;

            Main.GetWindowDispatcher().BeginInvoke(DispatcherPriority.Normal, new Action(() => MainWindow.GetInstance().Show()));
            //MainWindow.GetInstance().ShowWindow();
        }

        public void NewTextualClickEvent()
        {
            if (MainWindow.GetInstance() == null)
                return;

            //MainWindow.GetInstance().ShowWindow();
            //MainWindow.GetInstance().CreateMacroAsync(null, MacroType.PYTHON, "/");
        }

        public void NewVisualClickEvent()
        {
            if (MainWindow.GetInstance() == null)
                return;

            //MainWindow.GetInstance().ShowWindow();
            //MainWindow.GetInstance().CreateMacroAsync(null, MacroType.BLOCKLY, "/");
        }

        public void OpenMacroClickEvent()
        {
            if (MainWindow.GetInstance() == null)
                return;

            //MainWindow.GetInstance().ShowWindow();
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
            Main.GetWindowDispatcher().BeginInvoke(DispatcherPriority.Normal, new Action(() => MainWindow.GetInstance().TryFocus()));
        }

        private static void WindowShowEvent()
        {
            Main.GetWindowDispatcher().BeginInvoke(DispatcherPriority.Normal, new Action(() => MainWindow.GetInstance().ShowWindow()));
        }

        private static void WindowHideEvent()
        {
            Main.GetWindowDispatcher().BeginInvoke(DispatcherPriority.Normal, new Action(() => MainWindow.GetInstance().HideWindow()));
        }

        private static void DisplayOkMessage(string content, string title)
        {
            Main.GetWindowDispatcher().BeginInvoke(DispatcherPriority.Normal, new Action(() => MainWindow.GetInstance().DisplayOkMessage(content, title)));
        }

        private static void DisplayYesNoMessage(string content, string title, Action<bool> OnReturn)
        {
            Main.GetWindowDispatcher().BeginInvoke(DispatcherPriority.Normal, new Action(() => MainWindow.GetInstance().DisplayYesNoMessage(content, title, OnReturn)));
        }

        #endregion
    }
}
