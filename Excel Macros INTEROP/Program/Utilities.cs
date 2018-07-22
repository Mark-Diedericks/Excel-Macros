/*
 * Mark Diedericks
 * 22/07/2018
 * Version 1.0.3
 * Utility functions for the application
 */

using Excel_Macros_INTEROP.Engine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Macros_INTEROP
{
    public class Utilities
    {
        private static Utilities s_Instance;
        private static ExcelUtilities s_ExcelUtilities;
        private Dictionary<int, HighPrecisionTimer> m_DebugSessions;

        public Utilities()
        {
            s_Instance = this;
            s_ExcelUtilities = new ExcelUtilities();
            m_DebugSessions = new Dictionary<int, HighPrecisionTimer>();
        }

        public static Utilities GetInstance()
        {
            return s_Instance != null ? s_Instance : new Utilities();
        }

        public static void Instantiate()
        {
            new Utilities();
        }

        public static ExcelUtilities GetExcelUtilities()
        {
            return s_ExcelUtilities;
        }

        public static int BeginProfileSession()
        {
            int id = GetInstance().m_DebugSessions.Count;
            GetInstance().m_DebugSessions.Add(id, new HighPrecisionTimer());

            //Start debug timer
            GetInstance().m_DebugSessions[id].Start();

            return id;
        }

        public static void EndProfileSession(int id)
        {
            if (id == -1)
                return;

            if (!GetInstance().m_DebugSessions.ContainsKey(id))
                return;

            //Stop debug timer
            GetInstance().m_DebugSessions[id].Stop();

            GetInstance().m_DebugSessions.Remove(id);
        }

        public static double GetTimeIntervalMilli(int id)
        {
            if (id == -1)
                return 0.00;

            GetInstance().m_DebugSessions[id].Stop();
            double duration = GetInstance().m_DebugSessions[id].Duration; //Convert from milliseconds to milliseconds
            GetInstance().m_DebugSessions[id].Start();

            return duration;
        }

        public static double GetTimeIntervalSeconds(int id)
        {
            if (id == -1)
                return 0.00;

            GetInstance().m_DebugSessions[id].Stop();
            double duration = GetInstance().m_DebugSessions[id].Duration / 1000.0f; //Convert from milliseconds to seconds
            GetInstance().m_DebugSessions[id].Start();

            return duration;
        }

        //Get the excel application
        public Excel.Application GetApplication()
        {
            return Main.GetApplication();
        }

        //Get the active workbook
        public Excel.Workbook GetActiveWorkbook()
        {
            return GetApplication().ActiveWorkbook;
        }

        //Get the active worksheet
        public Excel.Worksheet GetActiveWorksheet()
        {
            return (Excel.Worksheet)GetActiveWorkbook().ActiveSheet;
        }
    }
}
