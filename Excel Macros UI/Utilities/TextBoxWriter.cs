/*
 * Mark Diedericks
 * 22/07/2018
 * Version 1.0.0
 * TextBox Text Writer
 */
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace Excel_Macros_UI.Utilities
{
    public class TextBoxWriter : TextWriter
    {
        public override Encoding Encoding
        {
            get
            {
                return Encoding.UTF8;
            }
        }

        private TextBox m_TextBox;

        public TextBoxWriter(TextBox textBox)
        {
            m_TextBox = textBox;
        }

        public override void Write(char value)
        {
            m_TextBox.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, (Action)(() => m_TextBox.Text += value.ToString()));
        }

        public override void Write(string value)
        {
            m_TextBox.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, (Action)(() => m_TextBox.Text += value.ToString()));
        }

        public override void Write(char[] buffer, int index, int count)
        {
            m_TextBox.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, (Action)(() => m_TextBox.Text += new string(buffer)));
        }

        public override void WriteLine(string value)
        {
            m_TextBox.Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Normal, (Action)(() => m_TextBox.Text += value.ToString() + '\n'));
        }
    }
}
