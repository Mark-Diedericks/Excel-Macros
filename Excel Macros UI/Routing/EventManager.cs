/*
 * Mark Diedericks
 * 02/08/2018
 * Version 1.0.9
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
using Excel_Macros_UI.ViewModel;
using Excel_Macros_INTEROP.Libraries;

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

        /// <summary>
        /// Instiantiation of EventManager
        /// </summary>
        private EventManager()
        {
            s_Instance = this;
            s_IsLoaded = false;
        }

        /// <summary>
        /// Gets instance of EventManager
        /// </summary>
        /// <returns></returns>
        public static EventManager GetInstance()
        {
            return s_Instance;
        }

        /// <summary>
        /// Returns whether or not the application has been loaded
        /// </summary>
        /// <returns>Whether or not the application is loaded</returns>
        public static bool IsLoaded()
        {
            return s_IsLoaded;
        }

        /// <summary>
        /// Returns whether or not the ribbon tab has been loaded
        /// </summary>
        /// <returns>Whether or not the application is loaded</returns>
        public static bool IsRibbonLoaded()
        {
            return s_IsRibbonLoaded;
        }

        /// <summary>
        /// Initializes the UI and Interop and binds events to the AddIn.
        /// </summary>
        /// <param name="application">Excel Application</param>
        /// <param name="dispatcher">Excel UI Dispatcher</param>
        /// <param name="RibbonMacros">A serialized list of ribbon accessible macros</param>
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
            }), RibbonMacros, Properties.Settings.Default.ActiveMacro, Properties.Settings.Default.IncludedLibraries);

            GetInstance().ShutdownEvent += () =>
            {
                MainWindow.GetInstance().Dispatcher.BeginInvoke(DispatcherPriority.Send, new Action(() =>
                {
                    if(Main.GetDeclaration(Main.GetActiveMacro()) != null)
                        Properties.Settings.Default.ActiveMacro = Main.GetDeclaration(Main.GetActiveMacro()).relativepath;
                    
                    Properties.Settings.Default.IncludedLibraries = Main.GetAssemblies().ToArray<AssemblyDeclaration>();

                    if (MainWindowViewModel.GetInstance() != null)
                    {
                        MainWindowViewModel.GetInstance().SaveAll();
                        List<DocumentViewModel> unsaved = MainWindowViewModel.GetInstance().DockManager.GetUnsavedDocuments();

                        if (unsaved.Count > 0)
                        {
                            bool save = DisplayYesNoMessageReturn("You have unsaved documents. Would you like to save them?", "Unsaved Documents");

                            if (save)
                                foreach (DocumentViewModel document in unsaved)
                                    if (document is TextualEditorViewModel)
                                        document.Save(null);
                        }
                    }

                    s_UIApp.Shutdown();
                }));
            };

            s_UIApp.Run();
        }

        #region UI Events

        /// <summary>
        /// Fires ThemeChanged event
        /// </summary>
        public static void ThemeChanged()
        {
            ThemeChangedEvent?.Invoke();
        }

        /// <summary>
        /// Fires DocumentChanged event
        /// </summary>
        /// <param name="document"></param>
        public static void DocumentChanged(DocumentViewModel document)
        {
            DocumentChangedEvent?.Invoke(document);
        }

        #endregion

        #region Main to UI to Excel Events

        /// <summary>
        /// Fowards event to Excel's InputBox -> Asynchronous
        /// </summary>
        /// <param name="message"></param>
        /// <param name="title"></param>
        /// <param name="def"></param>
        /// <param name="left"></param>
        /// <param name="top"></param>
        /// <param name="helpFile"></param>
        /// <param name="helpContextID"></param>
        /// <param name="type"></param>
        /// <param name="OnResult"></param>
        private static void EventManager_DisplayInputMessageEvent(string message, object title, object def, object left, object top, object helpFile, object helpContextID, object type, Action<object> OnResult)
        {
            GetInstance().DisplayInputMessageEvent?.Invoke(message, title, def, left, top, helpFile, helpContextID, type, OnResult);
        }

        /// <summary>
        /// Forwards event to Excel's InputBox -> Synchronous
        /// </summary>
        /// <param name="message"></param>
        /// <param name="title"></param>
        /// <param name="def"></param>
        /// <param name="left"></param>
        /// <param name="top"></param>
        /// <param name="helpFile"></param>
        /// <param name="helpContextID"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        private static object EventManager_DisplayInputMessageReturnEvent(string message, object title, object def, object left, object top, object helpFile, object helpContextID, object type)
        {
            return GetInstance().DisplayInputMessageReturnEvent?.Invoke(message, title, def, left, top, helpFile, helpContextID, type);
        }

        #endregion

        #region Excel to UI Events

        /// <summary>
        /// Fires Shutdown Event
        /// </summary>
        public void Shutdown()
        {
            GetInstance().ShutdownEvent?.Invoke();
        }

        /// <summary>
        /// Shows the main window
        /// </summary>
        public void MacroEditorClickEvent()
        {
            if (MainWindowViewModel.GetInstance() == null || MainWindow.GetInstance() == null)
                return;

            MainWindow.GetInstance().Dispatcher.Invoke(() =>
            {
                MainWindowViewModel.GetInstance().ShowWindow();
            });
        }

        /// <summary>
        /// Shows the main window and creates new textual macro
        /// </summary>
        public void NewTextualClickEvent()
        {
            if (MainWindowViewModel.GetInstance() == null || MainWindow.GetInstance() == null)
                return;

            MainWindow.GetInstance().Dispatcher.Invoke(() =>
            {
                MainWindowViewModel.GetInstance().ShowWindow();
                MainWindowViewModel.GetInstance().CreateMacroAsync(MacroType.PYTHON);
            });
        }

        /// <summary>
        /// Shows the main window and creates new visual macro
        /// </summary>
        public void NewVisualClickEvent()
        {
            if (MainWindowViewModel.GetInstance() == null || MainWindow.GetInstance() == null)
                return;

            MainWindow.GetInstance().Dispatcher.Invoke(() =>
            {
                MainWindowViewModel.GetInstance().ShowWindow();
                MainWindowViewModel.GetInstance().DockManager.ActiveContent = MainWindowViewModel.GetInstance().DockManager.VisualEditor;
            });
        }

        /// <summary>
        /// SHows the main window and prompts to import a macro
        /// </summary>
        public void OpenMacroClickEvent()
        {
            if (MainWindowViewModel.GetInstance() == null || MainWindow.GetInstance() == null)
                return;

            MainWindow.GetInstance().Dispatcher.Invoke(() =>
            {
                MainWindowViewModel.GetInstance().ShowWindow();
                MainWindowViewModel.GetInstance().ImportMacroAsync();
            });
        }

        /// <summary>
        /// Fires RibbonLoaded event
        /// </summary>
        public static void MacroRibbonLoaded()
        {
            s_IsRibbonLoaded = true;
            RibbonLoadedEvent?.Invoke();
        }

        #endregion

        #region Main to Ribbon Events

        /// <summary>
        /// Fires AddMacro event
        /// </summary>
        /// <param name="id"></param>
        /// <param name="macroName"></param>
        /// <param name="macroPath"></param>
        /// <param name="OnClick"></param>
        public void AddMacro(Guid id, string macroName, string macroPath, Action OnClick)
        {
            AddRibbonMacroEvent?.Invoke(id, macroName, macroPath, OnClick);
        }

        /// <summary>
        /// Fires RemoveMacro event
        /// </summary>
        /// <param name="id"></param>
        public void RemoveMacro(Guid id)
        {
            RemoveRibbonMacroEvent?.Invoke(id);
        }

        /// <summary>
        /// Fires RenameMacro event
        /// </summary>
        /// <param name="id"></param>
        /// <param name="macroName"></param>
        /// <param name="macroPath"></param>
        public void RenameMacro(Guid id, string macroName, string macroPath)
        {
            MacroDeclaration md = Main.GetDeclaration(id);
            RenameRibbonMacroEvent?.Invoke(id, macroName, macroPath);
        }

        /// <summary>
        /// Fires ApplicationLoaded event
        /// </summary>
        public static void LoadCompleted()
        {
            s_IsLoaded = true;
            ApplicationLoadedEvent?.Invoke();
        }

        #endregion

        #region Main to UI Events
        
        /// <summary>
        /// Focuses main window
        /// </summary>
        public static void WindowFocusEvent()
        {
            MainWindowViewModel.GetInstance().TryFocus();
        }

        /// <summary>
        /// Shows main window
        /// </summary>
        public static void WindowShowEvent()
        {
            MainWindowViewModel.GetInstance().ShowWindow();
        }

        /// <summary>
        /// Hides main window
        /// </summary>
        public static void WindowHideEvent()
        {
            MainWindowViewModel.GetInstance().HideWindow();
        }

        /// <summary>
        /// Displays OK message
        /// </summary>
        /// <param name="content">The message to be displayed</param>
        /// <param name="title">The message's header</param>
        public static void DisplayOkMessage(string content, string title)
        {
            MainWindowViewModel.GetInstance().DisplayOkMessage(content, title);
        }

        /// <summary>
        /// Displays yes/no message -> asynchronous
        /// </summary>
        /// <param name="content">The message to be displayed</param>
        /// <param name="title">The message's header</param>
        /// <param name="OnReturn">The Action, and bool representing the user's input, to be fires when the user returns input</param>
        public static void DisplayYesNoMessage(string content, string title, Action<bool> OnReturn)
        {
            MainWindowViewModel.GetInstance().DisplayYesNoMessage(content, title, OnReturn);
        }

        /// <summary>
        /// Displays yes/no message -> synchronous
        /// </summary>
        /// <param name="content">The message to be displayed</param>
        /// <param name="title">The message's header</param>
        /// <returns>Bool representing the user's input</returns>
        public static bool DisplayYesNoMessageReturn(string content, string title)
        {
            Task<bool> t = MainWindowViewModel.GetInstance().DisplayYesNoMessageReturn(content, title);
            t.Wait();
            return t.Result;
        }

        /// <summary>
        /// Displays yes/no message -> synchronous
        /// </summary>
        /// <param name="message">The message to be displayed</param>
        /// <param name="caption">The message's header</param>
        /// <param name="aux">The text in the 3rd button</param>
        /// <param name="OnReturn">The Action, and MessageDialogResult of the user's input, to be fired when the user provides input</param>
        public static void DisplayYesNoCancelMessage(string message, string caption, string aux, Action<MessageDialogResult> OnReturn)
        {
            MainWindowViewModel.GetInstance().DisplayYesNoCancelMessage(message, caption, aux, OnReturn);
        }
        
        /// <summary>
        /// Fires ClearAllIO event
        /// </summary>
        public static void ClearAllIO()
        {
            GetInstance().ClearAllIOEvent?.Invoke();
        }

        /// <summary>
        /// Fires IOChange event
        /// </summary>
        /// <param name="output">TextWriter for output stream</param>
        /// <param name="error">TextWriter for error stream</param>
        public static void ChangeIO(TextWriter output, TextWriter error)
        {
            GetInstance().IOChangedEvent?.Invoke(output, error);
        }

        #endregion
    }
}
