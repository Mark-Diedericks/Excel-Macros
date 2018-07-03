/*
 * Mark Diedericks
 * 17/06/2015
 * Version 1.0.0
 * Textual Editor UI Control
 */

using Excel_Macros_UI.Utilities;
using ICSharpCode.AvalonEdit.Highlighting;
using ICSharpCode.AvalonEdit.Highlighting.Xshd;
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
using System.Xml;

namespace Excel_Macros_UI.View
{
    /// <summary>
    /// Macro Editor Routed Commands
    /// </summary>
    public static class MacroEditorCommands
    {
        public static readonly RoutedCommand SaveMacro = new RoutedCommand();
        public static readonly RoutedCommand RunMacro = new RoutedCommand();
    }

    /// <summary>
    /// Interaction logic for TextualEditorView.xaml
    /// </summary>
    public partial class TextualEditorView : UserControl
    {
        public TextualEditorView()
        {
            InitializeComponent();

            //m_CodeEditor.SyntaxHighlighting = HighlightingLoader.Load(new XmlTextReader(SyntaxStyleLoader.GetStyleStream()), HighlightingManager.Instance);
            SyntaxStyleLoader.LoadColorValues();
            SyntaxStyleLoader.OnStyleChanged += delegate () 
            {
                m_CodeEditor.SyntaxHighlighting = HighlightingLoader.Load(new XmlTextReader(SyntaxStyleLoader.GetStyleStream()), HighlightingManager.Instance);
            };
        }

        #region Editor Event Callbacks

        private void CodeEditor_TextChanged(object sender, EventArgs e)
        {
            string src = m_CodeEditor.Text;
            //if (m_CurrentMacro != Guid.Empty)
            //    Program.Main.GetExcelDispatcher().Invoke(() => Program.Main.GetMacros()[m_CurrentMacro].SetSource(src));
        }

        private void SaveMacro_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void SaveMacro_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            //if (m_CurrentMacro != Guid.Empty)
            //{
            //    string src = m_CodeEditor.Text;
            //    Program.Main.GetExcelDispatcher().Invoke(() => Program.Main.GetMacros()[m_CurrentMacro].SetSource(src));
            //    Program.Main.GetExcelDispatcher().Invoke(() => Program.Main.GetMacros()[m_CurrentMacro].Save());
            //}

            e.Handled = true;
        }

        private void RunMacro_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        private void RunMacro_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            Excel_Macros_INTEROP.Engine.ExecutionEngine.GetDebugEngine().ExecuteMacro(m_CodeEditor.Text, null, true);

            e.Handled = true;
        }

        #endregion
    }
}
