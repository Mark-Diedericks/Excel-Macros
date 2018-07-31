/*
 * Mark Diedericks
 * 26/07/2018
 * Version 1.0.1
 * File Explorer UI Control
 */

using Excel_Macros_INTEROP;
using Excel_Macros_INTEROP.Macros;
using Excel_Macros_UI.Model;
using Excel_Macros_UI.ViewModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

namespace Excel_Macros_UI.View
{

    /// <summary>
    /// Interaction logic for ExplorerView.xaml
    /// </summary>
    public partial class ExplorerView : UserControl
    {
        
        public ExplorerView()
        {
            InitializeComponent();

            ThemeChanged();
            Routing.EventManager.ThemeChangedEvent += ThemeChanged;
        }
        
        private ResourceDictionary ThemeDictionary
        {
            get
            {
                return Resources.MergedDictionaries[1];
            }
        }

        private void ThemeChanged()
        {
            ThemeDictionary.MergedDictionaries.Clear();

            foreach (Uri uri in MainWindowViewModel.GetInstance().ActiveTheme.UriList)
                ThemeDictionary.MergedDictionaries.Add(new ResourceDictionary() { Source = uri });
        }

        private void ExplorerView_DataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            ((ExplorerViewModel)DataContext).FocusEvent += delegate () { tvMacroView.Focus(); };
        }

        private void tvMacroView_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            tvMacroView.ContextMenu = ((ExplorerViewModel)DataContext).CreateTreeViewContextMenu();
            tvMacroView.ContextMenu.IsOpen = true;
        }

        private void TreeViewItem_Selected(object sender, RoutedEventArgs e)
        {
            TreeViewItem item = sender as TreeViewItem;
            DisplayableTreeViewItem data = item.DataContext as DisplayableTreeViewItem;

            if (data != null)
                data.SelectedEvent?.Invoke(sender, e);
        }

        private void TreeViewItem_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            TreeViewItem item = sender as TreeViewItem;
            DisplayableTreeViewItem data = item.DataContext as DisplayableTreeViewItem;

            if (data != null)
                data.DoubleClickEvent?.Invoke(sender, e);
        }

        private void TreeViewItem_PreviewMouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {
            TreeViewItem item = sender as TreeViewItem;
            DisplayableTreeViewItem data = item.DataContext as DisplayableTreeViewItem;

            if (data != null)
                data.RightClickEvent?.Invoke(sender, e);
        }

        private void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            TextBox item = sender as TextBox;
            DisplayableTreeViewItem data = item.DataContext as DisplayableTreeViewItem;

            if (data != null)
                data.FocusLostEvent?.Invoke(sender, e);
        }

        private void TextBox_KeyUp(object sender, KeyEventArgs e)
        {
            TextBox item = sender as TextBox;
            DisplayableTreeViewItem data = item.DataContext as DisplayableTreeViewItem;

            if (data != null)
                data.KeyUpEvent?.Invoke(sender, e);
        }

        private void InputBox_IsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            UIElement uie = sender as UIElement;
            uie.Focus();
        }
    }
}
