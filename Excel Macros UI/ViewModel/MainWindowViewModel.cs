/*
 * Mark Diedericks
 * 24/07/2018
 * Version 1.0.4
 * Primary view model for handling main window's views
 */

using Excel_Macros_INTEROP;
using Excel_Macros_INTEROP.Macros;
using Excel_Macros_UI.Model;
using Excel_Macros_UI.Model.Base;
using Excel_Macros_UI.Routing;
using Excel_Macros_UI.Themes;
using Excel_Macros_UI.View;
using Excel_Macros_UI.ViewModel.Base;
using ICSharpCode.AvalonEdit.Document;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using System.Xml;
using Xceed.Wpf.AvalonDock.Controls;
using Xceed.Wpf.AvalonDock.Layout;
using Xceed.Wpf.AvalonDock.Layout.Serialization;

namespace Excel_Macros_UI.ViewModel
{
    public class MainWindowViewModel : Base.ViewModel, IThemeManager
    {
        private static MainWindowViewModel s_Instance;

        public MainWindowViewModel()
        {
            s_Instance = this;
            Model = new MainWindowModel();

            AddTheme(new LightTheme());
            AddTheme(new DarkTheme());

            SetTheme(Properties.Settings.Default.Theme);

            DockManager = new DockManagerViewModel(Properties.Settings.Default.OpenDocuments);

            SettingsMenu = new SettingsMenuViewModel();
        }

        public static MainWindowViewModel GetInstance()
        {
            return s_Instance;
        }

        private Dispatcher Dispatcher
        {
            get
            {
                return MainWindow.GetInstance().Dispatcher;
            }
        }

        private void InvokeWindow(Action a)
        {
            if (MainWindow.GetInstance() != null)
                MainWindow.GetInstance().Dispatcher.Invoke(a);
        }

        private void BeginInvokeWindow(Action a)
        {
            if (MainWindow.GetInstance() != null)
                MainWindow.GetInstance().Dispatcher.BeginInvoke(DispatcherPriority.Normal, a);
        }

        #region Model

        private MainWindowModel m_Model;
        public MainWindowModel Model
        {
            get
            {
                return m_Model;
            }
            set
            {
                if(m_Model != value)
                {
                    m_Model = value;
                    OnPropertyChanged(nameof(Model));
                }
            }
        }

        #endregion

        #region DockManager

        public DockManagerViewModel DockManager
        {
            get
            {
                return Model.DockManager;
            }
            set
            {
                if(Model.DockManager != value)
                {
                    Model.DockManager = value;
                    OnPropertyChanged(nameof(DockManager));
                }
            }
        }

        #endregion

        #region IsShown

        public bool IsShown
        {
            get
            {
                return Model.IsShown;
            }
            set
            {
                if (Model.IsShown != value)
                {
                    if (value)
                            BeginInvokeWindow(() => MainWindow.GetInstance().Show());
                    else
                        BeginInvokeWindow(() => MainWindow.GetInstance().Hide());

                    Model.IsShown = value;
                    OnPropertyChanged(nameof(IsShown));
                }
            }
        }

        #endregion

        #region IsFocused

        public bool IsFocused
        {
            get
            {
                return Model.IsFocused;
            }
            set
            {
                if (Model.IsFocused != value)
                {
                    if (value)

                        BeginInvokeWindow(() => MainWindow.GetInstance().Focus());

                    Model.IsShown = value;
                    OnPropertyChanged(nameof(IsFocused));
                }
            }
        }

        #endregion

        #region IsClosing

        public bool IsClosing
        {
            get
            {
                return Model.IsClosing;
            }
            set
            {
                if(Model.IsClosing != value)
                {
                    Model.IsClosing = value;
                    OnPropertyChanged(nameof(IsClosing));
                }
            }
        }

        #endregion

        #region IsExecuting

        public bool IsExecuting
        {
            get
            {
                return Model.IsExecuting;
            }
            set
            {
                if(Model.IsExecuting != value)
                {
                    Model.IsExecuting = value;
                    OnPropertyChanged(nameof(IsExecuting));
                    OnPropertyChanged(nameof(IsEditing));
                }
            }
        }

        #endregion

        #region IsEditing

        public bool IsEditing
        {
            get
            {
                return !IsExecuting;
            }
            set
            {
                if (IsExecuting == value)
                {
                    IsExecuting = !value;
                    OnPropertyChanged(nameof(IsEditing));
                }
            }
        }

        #endregion

        #region SettingsMenu

        public SettingsMenuViewModel SettingsMenu
        {
            get
            {
                return Model.SettingsMenu;
            }
            set
            {
                if(Model.SettingsMenu != value)
                {
                    Model.SettingsMenu = value;
                    OnPropertyChanged(nameof(SettingsMenu));
                }
            }
        }

        #endregion

        #region Themes

        public ObservableCollection<ITheme> Themes
        {
            get
            {
                return Model.Themes;
            }
            set
            {
                if(Model.Themes != value)
                {
                    Model.Themes = value;
                    OnPropertyChanged(nameof(Themes));
                }
            }
        }

        #endregion

        #region ActiveTheme

        public ITheme ActiveTheme
        {
            get
            {
                return Model.ActiveTheme;
            }
            set
            {
                if(Model.ActiveTheme != value)
                {
                    Model.ActiveTheme = value;
                    OnPropertyChanged(nameof(ActiveTheme));
                }
            }
        }

        #endregion

        #region DocumentContextMenu

        public ContextMenu DocumentContextMenu
        {
            get
            {
                return Model.DocumentContextMenu;
            }
            set
            {
                if(Model.DocumentContextMenu != value)
                {
                    Model.DocumentContextMenu = value;
                    Model.DocumentContextMenu.ContextMenuOpening += DocumentContextMenu_ContextMenuOpening;
                    OnPropertyChanged(nameof(DocumentContextMenu));
                }
            }
        }

        #endregion

        #region AnchorableContextMenu

        public ContextMenu AnchorableContextMenu
        {
            get
            {
                return Model.AnchorableContextMenu;
            }
            set
            {
                if (Model.AnchorableContextMenu != value)
                {
                    Model.AnchorableContextMenu = value;
                    Model.AnchorableContextMenu.ContextMenuOpening += AnchorableContextMenu_ContextMenuOpening;
                    OnPropertyChanged(nameof(AnchorableContextMenu));
                }
            }
        }

        #endregion

        #region AsyncExecution

        public bool AsyncExecution
        {
            get
            {
                return Model.AsyncExecution;
            }
            set
            {
                if(Model.AsyncExecution != value)
                {
                    Model.AsyncExecution = value;
                    OnPropertyChanged(nameof(AsyncExecution));
                }
            }
        }

        #endregion

        ///////////////////////////////
        ///////////////////////////////
        ///////////////////////////////

        #region Window Event Callbacks & Overrides

        public void OnClosing(CancelEventArgs e)
        {
            SaveAll();
            IsShown = false;

            e.Cancel = !IsClosing;
        }

        public void DockManagerLoaded()
        {
            LoadAvalonDockLayout();
        }

        public void DockManagerUnloaded()
        {
            SaveAvalonDockLayout();
        }

        #endregion

        #region Saving & Loading Settings

        public void SaveAll()
        {
            SaveAvalonDockLayout();
            Properties.Settings.Default.OpenDocuments = DockManager.GetVisibleDocuments();
            Properties.Settings.Default.Theme = ActiveTheme.Name;
            Properties.Settings.Default.Save();
        }

        private void SaveAvalonDockLayout()
        {
            if (MainWindow.GetInstance() != null)
                return;

            XmlLayoutSerializer serializer = new XmlLayoutSerializer(MainWindow.GetInstance().GetDockingManager());
            StringWriter stringWriter = new StringWriter();
            XmlWriter xmlWriter = XmlWriter.Create(stringWriter);

            serializer.Serialize(xmlWriter);

            xmlWriter.Flush();
            stringWriter.Flush();

            string layout = stringWriter.ToString();

            xmlWriter.Close();
            stringWriter.Close();

            Properties.Settings.Default.AvalonLayout = layout;
            Properties.Settings.Default.Save();
        }

        private void LoadAvalonDockLayout()
        {
            if (MainWindow.GetInstance() != null)
                return;

            XmlLayoutSerializer serializer = new XmlLayoutSerializer(MainWindow.GetInstance().GetDockingManager());
            serializer.LayoutSerializationCallback += (s, args) => { args.Content = args.Content; };

            string layout = Properties.Settings.Default.AvalonLayout;

            if (String.IsNullOrEmpty(layout.Trim()))
                return;

            StringReader stringReader = new StringReader(layout);
            XmlReader xmlReader = XmlReader.Create(stringReader);

            serializer.Deserialize(xmlReader);

            xmlReader.Close();
            stringReader.Close();
        }

        #endregion

        #region EventManager Event Function Callbacks

        public void ShowWindow()
        {
            IsShown = true;
            IsFocused = true;
            BeginInvokeWindow(() => MainWindow.GetInstance().Activate());
        }

        public void HideWindow()
        {
            SaveAll();
            IsShown = false;
        }

        public void CloseWindow()
        {
            IsClosing = true;
            BeginInvokeWindow(() => MainWindow.GetInstance().Close());
        }

        public void TryFocus()
        {
            IsFocused = true;
        }

        public void DisplayOkMessage(string message, string caption)
        {
            if (IsShown && MainWindow.GetInstance() != null)
                BeginInvokeWindow(async () => await MainWindow.GetInstance().ShowMessageAsync(caption, message, MessageDialogStyle.Affirmative, new MetroDialogSettings() { AffirmativeButtonText = "Ok" }));
            else
                System.Windows.Forms.MessageBox.Show(message, caption, System.Windows.Forms.MessageBoxButtons.OK);
        }

        public void DisplayYesNoMessage(string message, string caption, Action<bool> OnReturn)
        {
            if (IsShown && MainWindow.GetInstance() != null)
                BeginInvokeWindow(async () =>
                {
                    bool result = (await MainWindow.GetInstance().ShowMessageAsync(caption, message, MessageDialogStyle.AffirmativeAndNegative, new MetroDialogSettings() { AffirmativeButtonText = "Yes", NegativeButtonText = "No" })) == MessageDialogResult.Affirmative;
                    OnReturn?.Invoke(result);
                });
            else
                OnReturn?.Invoke(System.Windows.Forms.MessageBox.Show(message, caption, System.Windows.Forms.MessageBoxButtons.OK) == System.Windows.Forms.DialogResult.OK);
        }

        public void DisplayYesNoCancelMessage(string message, string caption, string aux, Action<MessageDialogResult> OnReturn)
        {
            if (IsShown && MainWindow.GetInstance() != null)
                BeginInvokeWindow(async () =>
                {
                    MessageDialogResult result = (await MainWindow.GetInstance().ShowMessageAsync(caption, message, MessageDialogStyle.AffirmativeAndNegativeAndSingleAuxiliary, new MetroDialogSettings() { AffirmativeButtonText = "Yes", NegativeButtonText = "Cancel", FirstAuxiliaryButtonText = aux }));
                    OnReturn?.Invoke(result);
                });
            else
                OnReturn?.Invoke(ConvertResult(System.Windows.Forms.MessageBox.Show(message, caption, System.Windows.Forms.MessageBoxButtons.YesNoCancel)));
        }

        private MessageDialogResult ConvertResult(System.Windows.Forms.DialogResult result)
        {
            switch (result)
            {
                case System.Windows.Forms.DialogResult.Yes:
                    return MessageDialogResult.Affirmative;
                case System.Windows.Forms.DialogResult.No:
                    return MessageDialogResult.Negative;
                case System.Windows.Forms.DialogResult.Cancel:
                    return MessageDialogResult.Canceled;
                default:
                    return MessageDialogResult.Canceled;
            }
        }

        public async Task<bool> DisplayYesNoMessageReturn(string message, string caption)
        {
            if (IsShown && MainWindow.GetInstance() != null)
                return await Dispatcher.Invoke(async () =>
                {
                    return (await MainWindow.GetInstance().ShowMessageAsync(caption, message, MessageDialogStyle.AffirmativeAndNegative)) == MessageDialogResult.Affirmative;
                });
            else
                return (System.Windows.Forms.MessageBox.Show(message, caption, System.Windows.Forms.MessageBoxButtons.OK) == System.Windows.Forms.DialogResult.OK);
        }

        #endregion

        #region Themes & Styles

        public bool AddTheme(ITheme theme)
        {
            if (Themes.Contains(theme))
                return false;

            Themes.Add(theme);
            return true;
        }

        public bool SetTheme(string name)
        {
            foreach (ITheme theme in Themes)
            {
                if (theme.Name.Trim().ToLower() == name.Trim().ToLower())
                {
                    ActiveTheme = theme;

                    InvokeWindow(() =>
                    {
                        MainWindow.GetInstance().ThemeDictionary.MergedDictionaries.Clear();

                        foreach (Uri uri in ActiveTheme.UriList)
                            MainWindow.GetInstance().ThemeDictionary.MergedDictionaries.Add(new ResourceDictionary() { Source = uri });
                    });

                    if (Properties.Settings.Default.Theme.Trim().ToLower() != name.Trim().ToLower())
                    {
                        Properties.Settings.Default.Theme = name.Trim();
                        Properties.Settings.Default.Save();
                    }
                    
                    Routing.EventManager.ThemeChanged();

                    InvokeWindow(() => ContextMenuThemeChange());

                    return true;
                }
            }

            return false;
        }

        #endregion

        #region Context Menus

        private void ContextMenuThemeChange()
        {
            if (DocumentContextMenu == null || AnchorableContextMenu == null)
                return;

            if (MainWindow.GetInstance() == null)
                return;

            Style ContextMenuStyle = MainWindow.GetInstance().GetResource("MetroContextMenuStyle") as Style;
            Style MenuItemStyle = MainWindow.GetInstance().GetResource("MetroMenuItemStyle") as Style;

            DocumentContextMenu.Resources.MergedDictionaries.Add(MainWindow.GetInstance().Resources);
            DocumentContextMenu.Style = ContextMenuStyle;
            DocumentContextMenu.ItemContainerStyle = MenuItemStyle;

            foreach (MenuItem item in DocumentContextMenu.Items)
                item.Style = MenuItemStyle;

            MainWindow.GetInstance().GetDockingManager().DocumentContextMenu = null;
            MainWindow.GetInstance().GetDockingManager().DocumentContextMenu = DocumentContextMenu;

            AnchorableContextMenu.Resources.MergedDictionaries.Add(MainWindow.GetInstance().Resources);
            AnchorableContextMenu.Style = ContextMenuStyle;
            AnchorableContextMenu.ItemContainerStyle = MenuItemStyle;

            foreach (MenuItem item in AnchorableContextMenu.Items)
                item.Style = MenuItemStyle;

            MainWindow.GetInstance().GetDockingManager().AnchorableContextMenu = null;
            MainWindow.GetInstance().GetDockingManager().AnchorableContextMenu = AnchorableContextMenu;
        }

        private void DocumentContextMenu_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            LayoutDocumentItem document = ((ContextMenu)sender).DataContext as LayoutDocumentItem;

            if (document != null)
            {
                DocumentViewModel model = document.Model as DocumentViewModel;

                if (model != null && model != MainWindow.GetInstance().GetDockingManager().ActiveContent)
                    MainWindow.GetInstance().GetDockingManager().ActiveContent = model;

                e.Handled = true;
                return;
            }

            e.Handled = false;
        }

        private void AnchorableContextMenu_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            LayoutAnchorableItem tool = ((ContextMenu)sender).DataContext as LayoutAnchorableItem;

            if (tool != null)
            {
                ToolViewModel model = tool.Model as ToolViewModel;

                if (model != null && model != MainWindow.GetInstance().GetDockingManager().ActiveContent)
                    MainWindow.GetInstance().GetDockingManager().ActiveContent = model;

                e.Handled = true;
                return;
            }

            e.Handled = false;
        }

        #endregion

        #region Active Documents
        
        private void ChangeActiveDocument(DocumentViewModel document)
        {
            if (DockManager.Documents.Contains(document))
                DockManager.ActiveContent = document;
        }

        #endregion

        #region Toolbar & MenuBar Event Callbacks

        #region NewClick

        private ICommand m_NewClick;
        public ICommand NewClick
        {
            get
            {
                if (m_NewClick == null)
                    m_NewClick = new RelayCommand(call => NewEvent());
                return m_NewClick;
            }
        }

        private void NewEvent()
        {
            DockManager.Explorer.CreateMacro(MacroType.PYTHON);
        }

        #endregion

        #region OpenClick

        private ICommand m_OpenClick;
        public ICommand OpenClick
        {
            get
            {
                if (m_OpenClick == null)
                    m_OpenClick = new RelayCommand(call => OpenEvent());
                return m_OpenClick;
            }
        }

        private void OpenEvent()
        {
            DockManager.Explorer.ImportMacro();
        }

        #endregion

        #region ExportClick

        private ICommand m_ExportClick;
        public ICommand ExportClick
        {
            get
            {
                if (m_ExportClick == null)
                    m_ExportClick = new RelayCommand(call => ExportEvent());
                return m_ExportClick;
            }
        }

        private void ExportEvent()
        {
            if (DockManager.ActiveDocument == null)
                return;

            IMacro macro = Main.GetMacro(DockManager.ActiveDocument.Macro);

            if (macro == null)
                return;

            macro.Export();
        }

        #endregion

        #region ExitClick

        private ICommand m_ExitClick;
        public ICommand ExitClick
        {
            get
            {
                if (m_ExitClick == null)
                    m_ExitClick = new RelayCommand(call => ExitEvent());
                return m_ExitClick;
            }
        }

        private void ExitEvent()
        {
            HideWindow();
        }

        #endregion

        #region SaveClick

        private ICommand m_SaveClick;
        public ICommand SaveClick
        {
            get
            {
                if (m_SaveClick == null)
                    m_SaveClick = new RelayCommand(call => SaveEvent());
                return m_SaveClick;
            }
        }

        private void SaveEvent()
        {
            if (DockManager.ActiveDocument == null)
                return;

            if (DockManager.ActiveDocument.SaveCommand.CanExecute(null))
                DockManager.ActiveDocument.SaveCommand.Execute(null);
        }

        #endregion

        #region SaveAllClick

        private ICommand m_SaveAllClick;
        public ICommand SaveAllClick
        {
            get
            {
                if (m_SaveAllClick == null)
                    m_SaveAllClick = new RelayCommand(call => SaveAllEvent());
                return m_SaveAllClick;
            }
        }

        private void SaveAllEvent()
        {
            foreach (DocumentViewModel document in DockManager.Documents)
            {
                if (document.SaveCommand.CanExecute(null))
                    document.SaveCommand.Execute(null);
            }
        }

        #endregion

        #region UndoClick

        private ICommand m_UndoClick;
        public ICommand UndoClick
        {
            get
            {
                if (m_UndoClick == null)
                    m_UndoClick = new RelayCommand(call => UndoEvent());
                return m_UndoClick;
            }
        }

        private void UndoEvent()
        {
            if (DockManager.ActiveDocument == null)
                return;

            if (DockManager.ActiveDocument.UndoCommand.CanExecute(null))
                DockManager.ActiveDocument.UndoCommand.Execute(null);
        }

        #endregion

        #region RedoClick

        private ICommand m_RedoClick;
        public ICommand RedoClick
        {
            get
            {
                if (m_RedoClick == null)
                    m_RedoClick = new RelayCommand(call => RedoEvent());
                return m_RedoClick;
            }
        }

        private void RedoEvent()
        {
            if (DockManager.ActiveDocument == null)
                return;

            if (DockManager.ActiveDocument.RedoCommand.CanExecute(null))
                DockManager.ActiveDocument.RedoCommand.Execute(null);
        }

        #endregion

        #region CopyClick

        private ICommand m_CopyClick;
        public ICommand CopyClick
        {
            get
            {
                if (m_CopyClick == null)
                    m_CopyClick = new RelayCommand(call => CopyEvent());
                return m_CopyClick;
            }
        }

        private void CopyEvent()
        {
            if (DockManager.ActiveDocument == null)
                return;

            if (DockManager.ActiveDocument.CopyCommand.CanExecute(null))
                DockManager.ActiveDocument.CopyCommand.Execute(null);
        }

        #endregion

        #region CutClick

        private ICommand m_CutClick;
        public ICommand CutClick
        {
            get
            {
                if (m_CutClick == null)
                    m_CutClick = new RelayCommand(call => CutEvent());
                return m_CutClick;
            }
        }

        private void CutEvent()
        {
            if (DockManager.ActiveDocument == null)
                return;

            if (DockManager.ActiveDocument.CutCommand.CanExecute(null))
                DockManager.ActiveDocument.CutCommand.Execute(null);
        }

        #endregion

        #region PasteClick

        private ICommand m_PasteClick;
        public ICommand PasteClick
        {
            get
            {
                if (m_PasteClick == null)
                    m_PasteClick = new RelayCommand(call => PasteEvent());
                return m_PasteClick;
            }
        }

        private void PasteEvent()
        {
            if (DockManager.ActiveDocument == null)
                return;

            if (DockManager.ActiveDocument.PasteCommand.CanExecute(null))
                DockManager.ActiveDocument.PasteCommand.Execute(null);
        }

        #endregion

        #region RunClick

        private ICommand m_RunClick;
        public ICommand RunClick
        {
            get
            {
                if (m_RunClick == null)
                    m_RunClick = new RelayCommand(call => RunEvent());
                return m_RunClick;
            }
        }

        private void RunEvent()
        {
            if (DockManager.ActiveDocument == null)
                return;

            if (DockManager.ActiveDocument.StartCommand.CanExecute(null))
            {
                /*btnStop.IsEnabled = true;
                btnStop.Visibility = Visibility.Visible;

                btnRun.IsEnabled = false;
                btnRun.Visibility = Visibility.Hidden;*/

                IsExecuting = true;

                DockManager.ActiveDocument.StartCommand.Execute(new Action(() =>
                {
                    /*btnStop.IsEnabled = false;
                    btnStop.Visibility = Visibility.Hidden;

                    btnRun.IsEnabled = true;
                    btnRun.Visibility = Visibility.Visible;*/

                    IsExecuting = false;
                }));
            }
        }

        #endregion

        #region StopClick

        private ICommand m_StopClick;
        public ICommand StopClick
        {
            get
            {
                if (m_StopClick == null)
                    m_StopClick = new RelayCommand(call => StopEvent());
                return m_StopClick;
            }
        }

        private void StopEvent()
        {
            if (DockManager.ActiveDocument == null)
                return;

            if (DockManager.ActiveDocument.StopCommand.CanExecute(null))
            {
                DockManager.ActiveDocument.StopCommand.Execute(new Action(() =>
                {
                    /*btnStop.IsEnabled = false;
                    btnStop.Visibility = Visibility.Hidden;

                    btnRun.IsEnabled = true;
                    btnRun.Visibility = Visibility.Visible;*/

                    IsExecuting = false;
                }));
            }
        }

        #endregion

        #endregion

        #region Menu Bar Event Callbacks

        #region StyleClick

        private ICommand m_StyleClick;
        public ICommand StyleClick
        {
            get
            {
                if (m_StyleClick == null)
                    m_StyleClick = new RelayCommand(call => StyleClickEvent());
                return m_StyleClick;
            }
        }

        private void StyleClickEvent()
        {
            SettingsMenu.IsOpen = true;
            SettingsMenu.StyleActive = true;
        }

        #endregion

        #region LibraryClick

        private ICommand m_LibraryClick;
        public ICommand LibraryClick
        {
            get
            {
                if (m_LibraryClick == null)
                    m_LibraryClick = new RelayCommand(call => LibraryClickEvent());
                return m_LibraryClick;
            }
        }

        private void LibraryClickEvent()
        {
            SettingsMenu.IsOpen = true;
            SettingsMenu.LibraryActive = true;
        }

        #endregion

        #region MacroClick

        private ICommand m_MacroClick;
        public ICommand MacroClick
        {
            get
            {
                if (m_MacroClick == null)
                    m_MacroClick = new RelayCommand(call => MacroClickEvent());
                return m_MacroClick;
            }
        }

        private void MacroClickEvent()
        {
            SettingsMenu.IsOpen = true;
            SettingsMenu.RibbonActive = true;
        }

        #endregion

        #region ToolboxClick

        private ICommand m_ToolboxClick;
        public ICommand ToolboxClick
        {
            get
            {
                if (m_ToolboxClick == null)
                    m_ToolboxClick = new RelayCommand(call => ToolboxClickEvent());
                return m_ToolboxClick;
            }
        }

        private void ToolboxClickEvent()
        {
            ShowAnchorable(DockManager.Toolbox.ContentId);
        }

        #endregion

        #region ExplorerClick

        private ICommand m_ExplorerClick;
        public ICommand ExplorerClick
        {
            get
            {
                if (m_ExplorerClick == null)
                    m_ExplorerClick = new RelayCommand(call => ExplorerClickEvent());
                return m_ExplorerClick;
            }
        }

        private void ExplorerClickEvent()
        {
            ShowAnchorable(DockManager.Explorer.ContentId);
        }

        #endregion

        #region ConsoleClick

        private ICommand m_ConsoleClick;
        public ICommand ConsoleClick
        {
            get
            {
                if (m_ConsoleClick == null)
                    m_ConsoleClick = new RelayCommand(call => ConsoleClickEvent());
                return m_ConsoleClick;
            }
        }

        private void ConsoleClickEvent()
        {
            ShowAnchorable(DockManager.Console.ContentId);
        }

        #endregion

        private void ShowAnchorable(string ContentId)
        {
            if (MainWindow.GetInstance() == null)
                return;

            foreach (ILayoutElement le in MainWindow.GetInstance().GetDockingManager().Layout.Children)
            {
                if (le is LayoutAnchorable)
                {
                    LayoutAnchorable la = le as LayoutAnchorable;

                    if (la.ContentId == ContentId)
                    {
                        la.Show();
                        return;
                    }
                }
            }
        }

        #endregion

        #region Macro Related Actions

        public void CloseMacro(Guid id)
        {
            DocumentViewModel dvm = DockManager.GetDocument(id);
            if (dvm != null)
                dvm.IsClosed = true;
        }

        public Guid CreateMacro(MacroType type, string relativepath)
        {
            return FileManager.CreateMacro(type, relativepath);
        }

        public void ImportMacro(string relativepath, Action<Guid> OnReturn)
        {
            FileManager.ImportMacro(relativepath, OnReturn);
        }

        public void RenameFolder(string olddir, string newdir)
        {
            foreach (Guid id in Main.RenameFolder(olddir, newdir))
            {
                DocumentViewModel dvm = DockManager.GetDocument(id);
                if (dvm != null)
                {
                    dvm.ToolTip = Main.GetDeclaration(id).relativepath;
                    dvm.ContentId = Main.GetDeclaration(id).relativepath;
                }
            }
        }

        public void RenameMacro(Guid id, string newName)
        {
            Main.RenameMacro(id, newName);

            DocumentViewModel dvm = DockManager.GetDocument(id);
            if (dvm != null)
            {
                dvm.Title = Main.GetDeclaration(id).name;
                dvm.ContentId = Main.GetDeclaration(id).relativepath;
            }
        }

        public void OpenMacroForEditing(Guid id)
        {
            if (id == Guid.Empty)
                return;

            DocumentViewModel dvm = DockManager.GetDocument(id);
            if (dvm != null)
            {
                DockManager.ActiveContent = dvm;
                return;
            }

            Main.SetActiveMacro(id);
            if (id != Guid.Empty)
            {
                DocumentModel model = DocumentModel.Create(id);

                if (model != null)
                {
                    DocumentViewModel viewModel = DocumentViewModel.Create(model);
                    DockManager.AddDocument(viewModel);
                    ChangeActiveDocument(viewModel);
                }
            }
        }

        public void ExecuteMacro(bool async)
        {
            AsyncExecution = async;

            if(RunClick.CanExecute(null))
                RunClick.Execute(null);
        }

        public void CreateMacroAsync(MacroType type)
        {
            if (NewClick.CanExecute(null))
                NewClick.Execute(null);
        }

        public void ImportMacroAsync()
        {
            if (OpenClick.CanExecute(null))
                OpenClick.Execute(null);
        }

        #endregion

    }
}
