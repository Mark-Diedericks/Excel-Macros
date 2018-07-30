/*
 * Mark Diedericks
 * 30/07/2018
 * Version 1.0.8
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
using Excel_Macros_UI.View;
using Excel_Macros_UI.Model;
using System.IO;
using MahApps.Metro.Controls.Dialogs;
using Excel_Macros_UI.ViewModel.Base;

namespace Excel_Macros_UI.Routing
{
    public class EventManager
    {
        public delegate void ClearIOEvent();
        public event ClearIOEvent ClearAllIOEvent;

        public delegate void IOChangeEvent(TextWriter output, TextWriter error);
        public event IOChangeEvent IOChangedEvent;

        public delegate void MacroAddEvent(Guid id, string macroName, string macroPath, Action macroClickEvent);
        public event MacroAddEvent AddRibbonMacroEvent;

        public delegate void MacroRemoveEvent(Guid id);
        public event MacroRemoveEvent RemoveRibbonMacroEvent;

        public delegate void MacroEditEvent(Guid id, string macroName, string macroPath);
        public event MacroEditEvent RenameRibbonMacroEvent;

        public delegate void LoadEvent();
        public static event LoadEvent ApplicationLoadedEvent;
        public static event LoadEvent RibbonLoadedEvent;
        private event LoadEvent ShutdownEvent;
        
        public delegate void InputMessageEvent(string message, object title, object def, object left, object top, object helpFile, object helpContextID, object type, Action<object> OnResult);
        public event InputMessageEvent DisplayInputMessageEvent;

        public delegate object InputMessageReturnEvent(string message, object title, object def, object left, object top, object helpFile, object helpContextID, object type);
        public event InputMessageReturnEvent DisplayInputMessageReturnEvent;

        public delegate void SetEnabled(bool enabled);
        public event SetEnabled SetExcelInteractiveEvent;
        public event SetEnabled FastWorkbookEvent;

        public delegate void SetEnabledWorksheet(Excel.Worksheet worksheet, bool enabled);
        public event SetEnabledWorksheet FastWorksheetEvent;

        public delegate void ThemeEvent();
        public static event ThemeEvent ThemeChangedEvent;

        public delegate void DocumentEvent(DocumentViewModel vm);
        public static event DocumentEvent DocumentChangedEvent;

        private static EventManager s_Instance;
        private static App s_UIApp;
        private static bool s_IsLoaded;
        private static bool s_IsRibbonLoaded;

        private EventManager()
        {
            s_Instance = this;
            s_IsLoaded = false;
        }

        public static EventManager GetInstance()
        {
            return s_Instance;
        }

        public static bool IsLoaded()
        {
            return s_IsLoaded;
        }

        public static bool IsRibbonLoaded()
        {
            return s_IsRibbonLoaded;
        }

        public static void CreateApplicationInstance(Excel.Application application, Dispatcher dispatcher, string RibbonMacros)
        {
            new EventManager();
            
            s_UIApp = new App();
            s_UIApp.InitializeComponent();

            RibbonLoadedEvent += Main.LoadRibbonMacros;

            Main.Initialize(application, dispatcher, new Action(() =>
            {
                Main.GetInstance().OnFocused += WindowFocusEvent;
                Main.GetInstance().OnShown += WindowShowEvent;
                Main.GetInstance().OnHidden += WindowHideEvent;

                MessageManager.GetInstance().DisplayOkMessageEvent += DisplayOkMessage;
                MessageManager.GetInstance().DisplayYesNoMessageEvent += DisplayYesNoMessage;
                MessageManager.GetInstance().DisplayYesNoMessageReturnEvent += DisplayYesNoMessageReturn;

                MessageManager.GetInstance().DisplayInputMessageEvent += EventManager_DisplayInputMessageEvent;
                MessageManager.GetInstance().DisplayInputMessageReturnEvent += EventManager_DisplayInputMessageReturnEvent;

                Excel_Macros_INTEROP.EventManager.GetInstance().ClearAllIOEvent += ClearAllIO;
                Excel_Macros_INTEROP.EventManager.GetInstance().AddRibbonMacroEvent += GetInstance().AddMacro;
                Excel_Macros_INTEROP.EventManager.GetInstance().RemoveRibbonMacroEvent += GetInstance().RemoveMacro;
                Excel_Macros_INTEROP.EventManager.GetInstance().RenameRibbonMacroEvent += GetInstance().RenameMacro;

                Excel_Macros_INTEROP.EventManager.GetInstance().SetExcelInteractive += (enabled) => 
                {
                    GetInstance().SetExcelInteractiveEvent?.Invoke(enabled);
                };

                Excel_Macros_INTEROP.EventManager.GetInstance().FastWorkbookEvent += (enabled) =>
                {
                    GetInstance().FastWorkbookEvent?.Invoke(enabled);
                };

                Excel_Macros_INTEROP.EventManager.GetInstance().FastWorksheetEvent += (worksheet, enabled) =>
                {
                    GetInstance().FastWorksheetEvent?.Invoke(worksheet, enabled);
                };

                GetInstance().IOChangedEvent += Main.SetIOSteams;

                if (s_IsRibbonLoaded)
                    Main.LoadRibbonMacros();

                LoadCompleted();
            }), RibbonMacros, Properties.Settings.Default.ActiveMacro);

            GetInstance().ShutdownEvent += () =>
            {
                MainWindow.GetInstance().Dispatcher.BeginInvoke(DispatcherPriority.Send, new Action(() =>
                {
                    if(Main.GetDeclaration(Main.GetActiveMacro()) != null)
                        Properties.Settings.Default.ActiveMacro = Main.GetDeclaration(Main.GetActiveMacro()).relativepath;

                    if (MainWindow.GetInstance() != null)
                        MainWindow.GetInstance().SaveAll();

                    s_UIApp.Shutdown();
                }));
            };

            s_UIApp.Run();
        }

        #region UI Events

        public static void ThemeChanged()
        {
            ThemeChangedEvent?.Invoke();
        }

        public static void DocumentChanged(DocumentViewModel document)
        {
            DocumentChangedEvent?.Invoke(document);
        }

        #endregion

        #region Main to UI to Excel Events

        private static void EventManager_DisplayInputMessageEvent(string message, object title, object def, object left, object top, object helpFile, object helpContextID, object type, Action<object> OnResult)
        {
            GetInstance().DisplayInputMessageEvent?.Invoke(message, title, def, left, top, helpFile, helpContextID, type, OnResult);
        }

        private static object EventManager_DisplayInputMessageReturnEvent(string message, object title, object def, object left, object top, object helpFile, object helpContextID, object type)
        {
            return GetInstance().DisplayInputMessageReturnEvent?.Invoke(message, title, def, left, top, helpFile, helpContextID, type);
        }

        #endregion

        #region Excel to UI Events

        public void Shutdown()
        {
            GetInstance().ShutdownEvent?.Invoke();
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
            MainWindow.GetInstance().CreateMacroAsync(MacroType.PYTHON);
        }

        public void NewVisualClickEvent()
        {
            if (MainWindow.GetInstance() == null)
                return;

            MainWindow.GetInstance().ShowWindow();
            MainWindow.GetInstance().CreateMacroAsync(MacroType.BLOCKLY);
        }

        public void OpenMacroClickEvent()
        {
            if (MainWindow.GetInstance() == null)
                return;

            MainWindow.GetInstance().ShowWindow();
            MainWindow.GetInstance().ImportMacroAsync();
        }

        public static void MacroRibbonLoaded()
        {
            s_IsRibbonLoaded = true;
            RibbonLoadedEvent?.Invoke();
        }

        #endregion

        #region Main to Ribbon Events

        public void AddMacro(Guid id, string macroName, string macroPath, Action OnClick)
        {
            AddRibbonMacroEvent?.Invoke(id, macroName, macroPath, OnClick);
        }

        public void RemoveMacro(Guid id)
        {
            RemoveRibbonMacroEvent?.Invoke(id);
        }

        public void RenameMacro(Guid id, string macroName, string macroPath)
        {
            MacroDeclaration md = Main.GetDeclaration(id);
            RenameRibbonMacroEvent?.Invoke(id, macroName, macroPath);
        }

        public static void LoadCompleted()
        {
            s_IsLoaded = true;
            ApplicationLoadedEvent?.Invoke();
        }

        #endregion

        #region Main to UI Events

        public static void WindowFocusEvent()
        {
            MainWindow.GetInstance().TryFocus();
        }

        public static void WindowShowEvent()
        {
            MainWindow.GetInstance().ShowWindow();
        }

        public static void WindowHideEvent()
        {
            MainWindow.GetInstance().HideWindow();
        }

        public static void DisplayOkMessage(string content, string title)
        {
            MainWindow.GetInstance().DisplayOkMessage(content, title);
        }

        public static void DisplayYesNoMessage(string content, string title, Action<bool> OnReturn)
        {
            MainWindow.GetInstance().DisplayYesNoMessage(content, title, OnReturn);
        }

        public static bool DisplayYesNoMessageReturn(string content, string title)
        {
            Task<bool> t = MainWindow.GetInstance().DisplayYesNoMessageReturn(content, title);
            t.Wait();
            return t.Result;
        }

        public static void DisplayYesNoCancelMessage(string message, string caption, string aux, Action<MessageDialogResult> OnReturn)
        {
            MainWindow.GetInstance().DisplayYesNoCancelMessage(message, caption, aux, OnReturn);
        }

        public static void ClearAllIO()
        {
            GetInstance().ClearAllIOEvent?.Invoke();
        }

        public static void ChangeIO(TextWriter output, TextWriter error)
        {
            GetInstance().IOChangedEvent?.Invoke(output, error);
        }

        #endregion
    }
}
