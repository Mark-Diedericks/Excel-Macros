﻿/*
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

        /// <summary>
        /// Private instatiation of Utilities 
        /// </summary>
        private Utilities()
        {
            s_Instance = this;
            s_ExcelUtilities = new ExcelUtilities();
            m_DebugSessions = new Dictionary<int, HighPrecisionTimer>();
        }

        /// <summary>
        /// Gets the instance of Utilities
        /// </summary>
        /// <returns>Utilities instance</returns>
        public static Utilities GetInstance()
        {
            return s_Instance != null ? s_Instance : new Utilities();
        }

        /// <summary>
        /// Public instatiation of Utilities
        /// </summary>
        public static void Instantiate()
        {
            new Utilities();
        }

        /// <summary>
        /// Gets instance of ExcelUtilities
        /// </summary>
        /// <returns></returns>
        public static ExcelUtilities GetExcelUtilities()
        {
            return s_ExcelUtilities;
        }

        /// <summary>
        /// Begins a new profiling session
        /// </summary>
        /// <returns>Profiling session identifier</returns>
        public static int BeginProfileSession()
        {
            int id = GetInstance().m_DebugSessions.Count;
            GetInstance().m_DebugSessions.Add(id, new HighPrecisionTimer());

            //Start debug timer
            GetInstance().m_DebugSessions[id].Start();

            return id;
        }

        /// <summary>
        /// Ends a profiling session
        /// </summary>
        /// <param name="id">Profiling session identifier</param>
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

        /// <summary>
        /// Gets the time interval of a profiling session
        /// </summary>
        /// <param name="id">Profiling session identifier</param>
        /// <returns>Time interval in milliseconds</returns>
        public static double GetTimeIntervalMilli(int id)
        {
            if (id == -1)
                return 0.00;

            GetInstance().m_DebugSessions[id].Stop();
            double duration = GetInstance().m_DebugSessions[id].Duration; //Convert from milliseconds to milliseconds
            GetInstance().m_DebugSessions[id].Start();

            return duration;
        }

        /// <summary>
        /// Gets the time interval of the profiling session
        /// </summary>
        /// <param name="id">Profiling session identifier</param>
        /// <returns>Time interval in seconds</returns>
        public static double GetTimeIntervalSeconds(int id)
        {
            if (id == -1)
                return 0.00;

            GetInstance().m_DebugSessions[id].Stop();
            double duration = GetInstance().m_DebugSessions[id].Duration / 1000.0f; //Convert from milliseconds to seconds
            GetInstance().m_DebugSessions[id].Start();

            return duration;
        }

        /// <summary>
        /// Get the excel application
        /// </summary>
        /// <returns>Excel Application</returns>
        public Excel.Application GetApplication()
        {
            return Main.GetApplication();
        }

        /// <summary>
        /// Get the active workbook
        /// </summary>
        /// <returns>Excel's ActiveWorkbook</returns>
        public Excel.Workbook GetActiveWorkbook()
        {
            return GetApplication().ActiveWorkbook;
        }

        /// <summary>
        /// Get the active worksheet
        /// </summary>
        /// <returns>Excel's ActiveWorksheet</returns>
        public Excel.Worksheet GetActiveWorksheet()
        {
            return (Excel.Worksheet)GetActiveWorkbook().ActiveSheet;
        }
    }
}
