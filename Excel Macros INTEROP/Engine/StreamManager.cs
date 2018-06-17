/*
 * Mark Diedericks
 * 09/06/2015
 * Version 1.0.0
 * Manages input, output and error streams of the execution engine
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_INTEROP.Engine
{
    public class StreamManager
    {
        #region Static Instancing

        private static StreamManager s_Instance;

        private static StreamManager GetInstance()
        {
            return s_Instance != null ? s_Instance : new StreamManager();
        }

        public static void Instantiate()
        {
            new StreamManager();
        }

        #endregion

        //Streams (in, out, err)
        private MemoryStream m_StreamIn;
        private MemoryStream m_StreamOut;
        private MemoryStream m_StreamErr;

        //Stream writers
        private StreamWriter m_WriteIn;
        private StreamWriter m_WriteOut;
        private StreamWriter m_WriteErr;

        //Stream readers
        private StreamReader m_ReadIn;
        private StreamReader m_ReadOut;
        private StreamReader m_ReadErr;

        private StreamManager()
        {
            s_Instance = this;

            //Create streams
            m_StreamIn = new MemoryStream();
            m_StreamOut = new MemoryStream();
            m_StreamErr = new MemoryStream();

            //Create stream writers
            m_WriteIn = new StreamWriter(m_StreamIn);
            m_WriteOut = new StreamWriter(m_StreamOut);
            m_WriteErr = new StreamWriter(m_StreamErr);

            //Create stream readers
            m_ReadIn = new StreamReader(m_StreamIn);
            m_ReadOut = new StreamReader(m_StreamOut);
            m_ReadErr = new StreamReader(m_StreamErr);
        }

        public static void ClearAllStreams()
        {
            ClearInputStream();
            ClearOutputStream();
            ClearErrorStream();
        }

        public static void ClearInputStream()
        {
            GetInstance().m_StreamIn.Flush();
        }

        public static void ClearOutputStream()
        {
            GetInstance().m_StreamOut.Flush();
        }

        public static void ClearErrorStream()
        {
            GetInstance().m_StreamErr.Flush();
        }

        public static MemoryStream GetInputStream()
        {
            return GetInstance().m_StreamIn;
        }

        public static MemoryStream GetOutputStream()
        {
            return GetInstance().m_StreamOut;
        }

        public static MemoryStream GetErrorStream()
        {
            return GetInstance().m_StreamErr;
        }

        public static StreamReader GetInputReader()
        {
            return GetInstance().m_ReadIn;
        }

        public static StreamReader GetOutputReader()
        {
            return GetInstance().m_ReadOut;
        }

        public static StreamReader GetErrorReader()
        {
            return GetInstance().m_ReadErr;
        }

        public static StreamWriter GetInputWriter()
        {
            return GetInstance().m_WriteIn;
        }

        public static StreamWriter GetOutputWriter()
        {
            return GetInstance().m_WriteOut;
        }

        public static StreamWriter GetErrorWriter()
        {
            return GetInstance().m_WriteErr;
        }

    }
}
