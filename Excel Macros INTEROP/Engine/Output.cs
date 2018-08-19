/*
 * Mark Diedericks
 * 19/08/2018
 * Version 1.0.0
 * Python output writer interfacing
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_INTEROP.Engine
{
    public class Output
    {
        private TextWriter m_TextWriter;

        public Output(TextWriter textWriter)
        {
            m_TextWriter = textWriter;
        }

        public void write(String str)
        {
            str = str.Replace("\n", Environment.NewLine);

            if (m_TextWriter != null)
                m_TextWriter.Write(str);
            else
                Console.Write(str);
        }

        public void writelines(String[] str)
        {
            foreach (String line in str)
            {
                if (m_TextWriter != null)
                    m_TextWriter.Write(str);
                else
                    Console.Write(str);
            }
        }

        public void flush()
        {
            if (m_TextWriter != null)
                m_TextWriter.Flush();
        }

        public void close()
        {
            if (m_TextWriter != null)
                m_TextWriter.Close();
        }
    }
}
