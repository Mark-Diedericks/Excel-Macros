/*
 * Mark Diedericks
 * 02/08/2018
 * Version 1.0.12
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
    public partial class MainWindow : MetroWindow
    {
        private static MainWindow s_Instance;

        public MainWindow()
        {
            s_Instance = this;

            InitializeComponent();

            ThemeManager.AddAccent("ExcelAccent", new Uri("pack://application:,,,/Excel Macros UI;component/Themes/ExcelAccent.xaml"));
            ThemeManager.ChangeAppStyle(this, ThemeManager.GetAccent("ExcelAccent"), ThemeManager.GetAppTheme("BaseLight"));

            this.DataContextChanged += MainWindow_DataContextChanged;
        }

        #region Events

        protected override void OnClosing(CancelEventArgs e)
        {
            ((MainWindowViewModel)DataContext).OnClosing(e);
        }
        
        private void DockManagerLoaded(object sender, RoutedEventArgs e)
        {
            ((MainWindowViewModel)DataContext).DockManagerLoaded();
        }

        private void DockManagerUnloaded(object sender, RoutedEventArgs e)
        {
            ((MainWindowViewModel)DataContext).DockManagerUnloaded();
        }

        private void MainWindow_DataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            ((MainWindowViewModel)DataContext).DocumentContextMenu = DockingManager_DockManager.DocumentContextMenu;
            ((MainWindowViewModel)DataContext).AnchorableContextMenu = DockingManager_DockManager.AnchorableContextMenu;
        }

        #endregion

        #region Getters

        public static MainWindow GetInstance()
        {
            return s_Instance;
        }

        public Xceed.Wpf.AvalonDock.DockingManager GetDockingManager()
        {
            return DockingManager_DockManager;
        }

        public ResourceDictionary ThemeDictionary
        {
            get
            {
                return Resources.MergedDictionaries[1];
            }
        }

        public object GetResource(string resource)
        {
            return Resources[resource];
        }

        public ResourceDictionary GetResources()
        {
            return Resources;
        }

        #endregion
    }
}
