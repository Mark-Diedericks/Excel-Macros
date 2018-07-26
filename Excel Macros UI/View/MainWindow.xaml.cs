/*
 * Mark Diedericks
 * 24/07/2018
 * Version 1.0.7
 * The main window, hosting all the UI
 */

using Excel_Macros_INTEROP;
using Excel_Macros_INTEROP.Macros;
using Excel_Macros_UI.Model;
using Excel_Macros_UI.Model.Base;
using Excel_Macros_UI.Themes;
using Excel_Macros_UI.Utilities;
using Excel_Macros_UI.ViewModel;
using Excel_Macros_UI.ViewModel.Base;
using MahApps.Metro;
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
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using System.Xml;
using Xceed.Wpf.AvalonDock.Controls;
using Xceed.Wpf.AvalonDock.Layout;
using Xceed.Wpf.AvalonDock.Layout.Serialization;

namespace Excel_Macros_UI.View
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : MetroWindow, IThemeManager
    {
        public delegate void ThemeEvent();
        public static event ThemeEvent ThemeChanged;

        public delegate void DocumentEvent(DocumentViewModel vm);
        public static event DocumentEvent DocumentChangedEvent;

        private static MainWindow s_Instance;
        private bool m_IsClosing;

        public MainWindow()
        {
            InitializeComponent();

            s_Instance = this;
            m_IsClosing = false;

            ThemeManager.AddAccent("ExcelAccent", new Uri("pack://application:,,,/Excel Macros UI;component/Themes/ExcelAccent.xaml"));
            ThemeManager.ChangeAppStyle(this, ThemeManager.GetAccent("ExcelAccent"), ThemeManager.GetAppTheme("BaseLight"));
            
            DockingManager_DockManager.DocumentContextMenu.ContextMenuOpening += DocumentContextMenu_ContextMenuOpening;
            DockingManager_DockManager.AnchorableContextMenu.ContextMenuOpening += AnchorableContextMenu_ContextMenuOpening;

            Themes = new ObservableCollection<ITheme>();
            
            AddTheme(new LightTheme());
            AddTheme(new DarkTheme());

            SetTheme(Properties.Settings.Default.Theme);

            flyoutSettings.DataContext = new SettingsMenuViewModel();
            flyoutSettings.CreateSettingsMenu();
        }

        public static MainWindow GetInstance()
        {
            return s_Instance;
        }

        #region Window Event Callbacks & Overrides

        protected override void OnClosing(CancelEventArgs e)
        {
            SaveAll();
            this.Hide();

            e.Cancel = !m_IsClosing;
        }

        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            
        }

        private void MetroWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {

        }

        private void MetroWindow_SizeChanged(object sender, SizeChangedEventArgs e)
        {

        }

        private void DockingManager_DockManager_Loaded(object sender, RoutedEventArgs e)
        {
            LoadAvalonDockLayout();

            if (DockingManager_DockManager.ActiveContent is DocumentViewModel)
                ActiveDocument = DockingManager_DockManager.ActiveContent as DocumentViewModel;
        }

        private void DockingManager_DockManager_Unloaded(object sender, RoutedEventArgs e)
        {
            SaveAvalonDockLayout();
        }

        #endregion

        #region Saving & Loading Settings

        public void SaveAll()
        {
            Dispatcher.BeginInvoke(DispatcherPriority.Send, new Action(delegate ()
            {
                SaveAvalonDockLayout();
                Properties.Settings.Default.OpenDocuments = ((DockManagerViewModel)DockingManager_DockManager.DataContext).GetVisibleDocuments();
                Properties.Settings.Default.Theme = ActiveTheme.Name;
                Properties.Settings.Default.Save();
            }));
        }

        private void SaveAvalonDockLayout()
        {
            XmlLayoutSerializer serializer = new XmlLayoutSerializer(DockingManager_DockManager);
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
            XmlLayoutSerializer serializer = new XmlLayoutSerializer(DockingManager_DockManager);
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
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
            {
                Show();
                Focus();
                Activate();
            }));
        }

        public void HideWindow()
        {
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
            {
                SaveAll();
                Close();
            }));
        }

        public void CloseWindow()
        {
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(delegate ()
            {
                base.Close();
                Dispatcher.BeginInvokeShutdown(DispatcherPriority.Normal);
            }));
        }

        public void TryFocus()
        {
            Dispatcher.BeginInvoke(DispatcherPriority.ContextIdle, new Action(delegate () { Focus(); }));
        }

        public void DisplayOkMessage(string message, string caption)
        {
            if (IsVisible)
                Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(async () => { await this.ShowMessageAsync(caption, message, MessageDialogStyle.Affirmative); }));
            else
                System.Windows.Forms.MessageBox.Show(message, caption, System.Windows.Forms.MessageBoxButtons.OK);
        }

        public void DisplayYesNoMessage(string message, string caption, Action<bool> OnReturn)
        {
            if (IsVisible)
                Dispatcher.Invoke(async () =>
                {
                    bool result = (await this.ShowMessageAsync(caption, message, MessageDialogStyle.AffirmativeAndNegative)) == MessageDialogResult.Affirmative;
                    OnReturn?.Invoke(result);
                });
            else
                OnReturn?.Invoke(System.Windows.Forms.MessageBox.Show(message, caption, System.Windows.Forms.MessageBoxButtons.OK) == System.Windows.Forms.DialogResult.OK);
        }

        public async Task<bool> DisplayYesNoMessageReturn(string message, string caption)
        {
            if (IsVisible)
                return await Dispatcher.Invoke(async () =>
                {
                    return (await this.ShowMessageAsync(caption, message, MessageDialogStyle.AffirmativeAndNegative)) == MessageDialogResult.Affirmative;
                });
            else
                return (System.Windows.Forms.MessageBox.Show(message, caption, System.Windows.Forms.MessageBoxButtons.OK) == System.Windows.Forms.DialogResult.OK);
        }

        #endregion

        #region IThemeManager

        public ResourceDictionary ThemeDictionary
        {
            get
            {
                return Resources.MergedDictionaries[1];
            }
        }

        public ObservableCollection<ITheme> Themes { get; internal set; }

        public ITheme ActiveTheme { get; internal set; }

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

                    ThemeDictionary.MergedDictionaries.Clear();
                    flyoutSettings.ThemeDictionary.MergedDictionaries.Clear();

                    foreach (Uri uri in ActiveTheme.UriList)
                    {
                        ThemeDictionary.MergedDictionaries.Add(new ResourceDictionary() { Source = uri });
                        flyoutSettings.ThemeDictionary.MergedDictionaries.Add(new ResourceDictionary() { Source = uri });
                    }

                    if(Properties.Settings.Default.Theme.Trim().ToLower() != name.Trim().ToLower())
                    {
                        Properties.Settings.Default.Theme = name.Trim();
                        Properties.Settings.Default.Save();
                    }

                    ThemeChanged?.Invoke();
                    flyoutSettings.Theme = FlyoutTheme.Accent;
                    flyoutSettings.Theme = FlyoutTheme.Adapt;

                    //DockingManager_DockManager.DocumentContextMenu = null;
                    //DockingManager_DockManager.DocumentContextMenu = CreateDocumentContextMenu();

                    //DockingManager_DockManager.AnchorableContextMenu = null;
                    //DockingManager_DockManager.AnchorableContextMenu = CreateAnchorableContextMenu();

                    return true;
                }
            }

            return false;
        }
        #endregion

        #region Context Menus

        private ContextMenu CreateDocumentContextMenu()
        {
            return FindResource("DocumentContextMenu") as ContextMenu;
        }

        private ContextMenu CreateAnchorableContextMenu()
        {
            return FindResource("AnchorableContextMenu") as ContextMenu;
        }

        private void DocumentContextMenu_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            LayoutDocumentItem document = ((ContextMenu)sender).DataContext as LayoutDocumentItem;

            if(document != null)
            {
                DocumentViewModel model = document.Model as DocumentViewModel;

                if (model != null && model != DockingManager_DockManager.ActiveContent)
                    DockingManager_DockManager.ActiveContent = model;

                /*ContextMenu cm = CreateDocumentContextMenu();
                cm.DataContext = document;

                ((ContextMenu)sender).IsOpen = false;
                cm.IsOpen = true;

                e.Handled = true;
                return;*/
            }

            e.Handled = false;
        }

        private void AnchorableContextMenu_ContextMenuOpening(object sender, ContextMenuEventArgs e)
        {
            LayoutAnchorableItem tool = ((ContextMenu)sender).DataContext as LayoutAnchorableItem;

            if (tool != null)
            {
                ToolViewModel model = tool.Model as ToolViewModel;

                if (model != null && model != DockingManager_DockManager.ActiveContent)
                    DockingManager_DockManager.ActiveContent = model;

                /*ContextMenu cm = CreateDocumentContextMenu();
                cm.DataContext = tool;
                
                ((ContextMenu)sender).IsOpen = false;
                cm.IsOpen = true;

                e.Handled = true;
                return;*/
            }
            
            e.Handled = false;
        }

        #endregion

        #region Active Documents

        private DocumentViewModel m_ActiveDocument;
        public DocumentViewModel ActiveDocument
        {
            get
            {
                return m_ActiveDocument;
            }
            set
            {
                m_ActiveDocument = value;
            }
        }

        private void DockingManager_DockManager_ActiveContentChanged(object sender, EventArgs e)
        {
            if (DockingManager_DockManager.ActiveContent is DocumentViewModel)
            {
                ActiveDocument = DockingManager_DockManager.ActiveContent as DocumentViewModel;
                DocumentChangedEvent?.Invoke(ActiveDocument);
            }
        }

        private void ChangeActiveDocument(DocumentViewModel document)
        {
            if(((DockManagerViewModel)DockingManager_DockManager.DataContext).Documents.Contains(document))
                DockingManager_DockManager.ActiveContent = document;
        }

        #endregion

        #region Toolbar Event Callbacks

        public bool AsyncExecution
        {
            get
            {
                return spltBtnExecutionType.SelectedIndex == 0;
            }

            set
            {
                spltBtnExecutionType.SelectedIndex = value ?  0 : 1;
            }
        }

        private void btnNew_Click(object sender, RoutedEventArgs e)
        {
            //FileManager.CreateMacro(Excel_Macros_INTEROP.Macros.MacroType.PYTHON, "");
            throw new NotImplementedException();
        }

        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            //FileManager.ImportMacro("", null);
            throw new NotImplementedException();
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            //if (ActiveDocument == null)
            //    return;
            //
            //if (ActiveDocument.SaveCommand.CanExecute(e))
            //    ActiveDocument.SaveCommand.Execute(e);
            throw new NotImplementedException();
        }

        private void btnExit_Click(object sender, RoutedEventArgs e)
        {
            HideWindow();
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (ActiveDocument == null)
                return;

            if (ActiveDocument.SaveCommand.CanExecute(null))
                ActiveDocument.SaveCommand.Execute(null);
        }

        private void btnSaveAll_Click(object sender, RoutedEventArgs e)
        {
            if (DockingManager_DockManager.DataContext == null)
                return;

            if (!(DockingManager_DockManager.DataContext is DockManagerViewModel))
                return;

            DockManagerViewModel dmvm = DockingManager_DockManager.DataContext as DockManagerViewModel;

            foreach(DocumentViewModel document in dmvm.Documents)
            {
                if (document.SaveCommand.CanExecute(null))
                    document.SaveCommand.Execute(null);
            }
        }

        private void btnUndo_Click(object sender, RoutedEventArgs e)
        {
            if (ActiveDocument == null)
                return;

            if (ActiveDocument.UndoCommand.CanExecute(null))
                ActiveDocument.UndoCommand.Execute(null);
        }

        private void btnRedo_Click(object sender, RoutedEventArgs e)
        {
            if (ActiveDocument == null)
                return;

            if (ActiveDocument.RedoCommand.CanExecute(null))
                ActiveDocument.RedoCommand.Execute(null);
        }

        private void btnCopy_Click(object sender, RoutedEventArgs e)
        {
            if (ActiveDocument == null)
                return;

            if (ActiveDocument.CopyCommand.CanExecute(null))
                ActiveDocument.CopyCommand.Execute(null);
        }

        private void btnCut_Click(object sender, RoutedEventArgs e)
        {
            if (ActiveDocument == null)
                return;

            if (ActiveDocument.CutCommand.CanExecute(null))
                ActiveDocument.CutCommand.Execute(null);
        }

        private void btnPaste_Click(object sender, RoutedEventArgs e)
        {
            if (ActiveDocument == null)
                return;

            if (ActiveDocument.PasteCommand.CanExecute(null))
                ActiveDocument.PasteCommand.Execute(null);
        }

        private void btnRun_Click(object sender, RoutedEventArgs e)
        {
            if (ActiveDocument == null)
                return;

            if (ActiveDocument.StartCommand.CanExecute(null))
            {
                btnStop.IsEnabled = true;
                btnStop.Visibility = Visibility.Visible;

                btnRun.IsEnabled = false;
                btnRun.Visibility = Visibility.Hidden;

                ActiveDocument.StartCommand.Execute(new Action(() =>
                {
                    GetInstance().Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() =>
                    {
                        btnStop.IsEnabled = false;
                        btnStop.Visibility = Visibility.Hidden;

                        btnRun.IsEnabled = true;
                        btnRun.Visibility = Visibility.Visible;
                    }));
                }));
            }
        }

        private void btnStop_Click(object sender, RoutedEventArgs e)
        {
            if (ActiveDocument == null)
                return;

            if (ActiveDocument.StopCommand.CanExecute(null))
            {
                ActiveDocument.StopCommand.Execute(new Action(() =>
                {
                    btnStop.IsEnabled = false;
                    btnStop.Visibility = Visibility.Hidden;

                    btnRun.IsEnabled = true;
                    btnRun.Visibility = Visibility.Visible;
                }));
            }
        }

        #endregion

        #region Menu Items
        
        private void Options_Click(object sender, RoutedEventArgs e)
        {
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() =>
            {
                flyoutSettings.IsOpen = true;
                flyoutSettings.SetActiveSettingsPage(SettingsMenuView.SettingsPage.Style);
            }));
        }

        private void Style_Click(object sender, RoutedEventArgs e)
        {
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() =>
            {
                flyoutSettings.IsOpen = true;
                flyoutSettings.SetActiveSettingsPage(SettingsMenuView.SettingsPage.Style);
            }));
        }

        private void Libraries_Click(object sender, RoutedEventArgs e)
        {
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() =>
            {
                flyoutSettings.IsOpen = true;
                flyoutSettings.SetActiveSettingsPage(SettingsMenuView.SettingsPage.Libraries);
            }));
        }

        private void Macros_Click(object sender, RoutedEventArgs e)
        {
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() =>
            {
                flyoutSettings.IsOpen = true;
                flyoutSettings.SetActiveSettingsPage(SettingsMenuView.SettingsPage.Macro);
            }));
        }

        private void Toolbox_Click(object sender, RoutedEventArgs e)
        {
            ShowAnchorable(((DockManagerViewModel)DockingManager_DockManager.DataContext).Toolbox.ContentId);
        }

        private void Explorer_Click(object sender, RoutedEventArgs e)
        {
            ShowAnchorable(((DockManagerViewModel)DockingManager_DockManager.DataContext).Explorer.ContentId);
        }

        private void Console_Click(object sender, RoutedEventArgs e)
        {
            ShowAnchorable(((DockManagerViewModel)DockingManager_DockManager.DataContext).Console.ContentId);
        }

        private void ShowAnchorable(string ContentId)
        {
            foreach (ILayoutElement le in DockingManager_DockManager.Layout.Children)
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
        
        public Guid CreateMacro(MacroType type, string relativepath)
        {
            return FileManager.CreateMacro(type, relativepath);
        }

        public void ImportMacro(string relativepath, Action<Guid> OnReturn)
        {
            FileManager.ImportMacro(relativepath, OnReturn);
        }

        public void OpenMacroForEditing(Guid id)
        {
            if (id == Guid.Empty)
                return;

            Dispatcher.Invoke(() => Main.SetActiveMacro(id));
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() => 
            {
                if (id != Guid.Empty)
                {
                    DocumentModel model = DocumentModel.Create(id);

                    if (model != null)
                    {
                        DocumentViewModel viewModel = DocumentViewModel.Create(model);
                        ((DockManagerViewModel)DockingManager_DockManager.DataContext).AddDocument(viewModel);
                        ChangeActiveDocument(viewModel);
                    }
                }
            }));
        }

        public void ExecuteMacro(bool async)
        {
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() =>
            {
                AsyncExecution = async;
                btnRun_Click(null, null);
            }));
        }

        /*public void CreateMacroAsync(TreeViewItem parent, MacroType type, string root)
        {
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() => m_MacroExplorer.CreateMacro(parent, type, root)));
        }

        public void ImportMacroAsync(TreeViewItem parent, string root)
        {
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() => m_MacroExplorer.ImportMacro(parent, root)));
        }

        public void ExecuteMacro(bool async)
        {
            Dispatcher.BeginInvoke(DispatcherPriority.Normal, new Action(() => m_MacroEditor.BeginExecution(async)));
        }*/

        #endregion
    }
}
