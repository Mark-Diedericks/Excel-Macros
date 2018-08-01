/*
 * Mark Diedericks
 * 17/06/2018
 * Version 1.0.0
 * Visual Editor UI Control
 */
 
using Excel_Macros_UI.ViewModel;
using System;
using System.Collections.Generic;
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

namespace Excel_Macros_UI.View
{
    /// <summary>
    /// Interaction logic for VisualEditorView.xaml
    /// </summary>
    public partial class VisualEditorView : UserControl
    {

        public VisualEditorView()
        {
            InitializeComponent();

            wbBlockly.Source = new Uri("pack://siteoforigin:,,,/Resources/BlocklyHost.html", UriKind.RelativeOrAbsolute);
            wbBlockly.LostFocus += (s, e) => { MainWindowViewModel.GetInstance().TryFocus(); };
        }

        private void VisualEditorView_DataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            VisualEditorViewModel vm = DataContext as VisualEditorViewModel;

            if (vm == null)
                return;

            vm.InvokeEngine += InvokeEngine;
        }

        private string InvokeEngine()
        {
            return wbBlockly.InvokeScript("showCode", new object[] { }).ToString();
        }
    }
}
