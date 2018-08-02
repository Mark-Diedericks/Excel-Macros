/*
 * Mark Diedericks
 * 02/08/2018
 * Version 1.0.3
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
        /// <summary>
        /// Instantiation of VisualEditorView
        /// </summary>
        public VisualEditorView()
        {
            InitializeComponent();

            wbBlockly.LoadCompleted += (s, e) => { SetSize(); };
            wbBlockly.Source = new Uri("pack://siteoforigin:,,,/Resources/BlocklyHost.html", UriKind.RelativeOrAbsolute);
            wbBlockly.LostFocus += (s, e) => { MainWindowViewModel.GetInstance().TryFocus(); };
        }

        /// <summary>
        /// DataContextChanged event callback, binds events
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void VisualEditorView_DataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            VisualEditorViewModel vm = DataContext as VisualEditorViewModel;

            if (vm == null)
                return;

            vm.InvokeEngine += InvokeEngine;
        }

        /// <summary>
        /// Invokes the Blockly UI to resize
        /// </summary>
        private void SetSize()
        {
            if (!wbBlockly.IsLoaded)
                return;

            wbBlockly.InvokeScript("resize", new object[] { });
            Panel.SetZIndex(wbBlockly, 1);
        }

        /// <summary>
        /// Invokes the Blockly to produce python code
        /// </summary>
        /// <returns>Source code (python)</returns>
        private string InvokeEngine()
        {
            return wbBlockly.InvokeScript("showCode", new object[] { }).ToString();
        }

        /// <summary>
        /// SizeChanged event callback, resizes Blockly
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void wbBlockly_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            SetSize();
        }
    }
}
