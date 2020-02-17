using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_INTEROP.Engine
{
    class PyOutput
    {
        private TextWriter myWriter = null;

        // ctor
        public PyOutput(TextWriter writer)
        {
            myWriter = writer;
        }

        public void write(String str)
        {
            str = str.Replace("\n", Environment.NewLine);
            if (myWriter != null)
            {
                myWriter.Write(str);
            }
            else
            {
                Console.Write(str);
            }
        }

        public void writelines(String[] str)
        {
            foreach (String line in str)
            {
                if (myWriter != null)
                {
                    myWriter.Write(str);
                }
                else
                {
                    Console.Write(str);
                }
            }
        }

        public void flush()
        {
            if (myWriter != null)
            {
                myWriter.Flush();
            }
        }

        public void close()
        {
            if (myWriter != null)
            {
                myWriter.Close();
            }
        }
    }
}
