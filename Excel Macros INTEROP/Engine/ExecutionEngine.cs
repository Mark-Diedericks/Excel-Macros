/*
 * Mark Diedericks
 * 02/08/2018
 * Version 1.0.10
 * Manages execution of python code
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Scripting.Hosting;
using System.Windows.Threading;
using System.ComponentModel;
using Python.Runtime;
using System.IO;

namespace Excel_Macros_INTEROP.Engine
{
    public class ExecutionEngine
    {
        #region Static Initializaton

        /// <summary>
        /// Intialize static instances of the ExecutionEngine
        /// </summary>
        public static void Initialize()
        {
            s_Engine = new ExecutionEngine();
        }

        private static ExecutionEngine s_Engine;

        /// <summary>
        /// Get instance of the Debug Execution Engine
        /// </summary>
        /// <returns>Execution Engine</returns>
        public static ExecutionEngine GetEngine()
        {
            if (s_Engine == null)
                Initialize();

            return s_Engine;
        }

        #endregion

        #region Instanced Initializaton
        
        private BackgroundWorker m_BackgroundWorker;
        private bool m_IsExecuting;

        /// <summary>
        /// Private Initialization of Exeuction Engine instance.
        /// </summary>
        private ExecutionEngine()
        {

#warning NOT COMPLETED DIRECTORIES

            string envPythonHome = @"E:\Mark Diedericks\Documents\Visual Studio 2017\Projects\Excel Macros\Dependencies\Python 27\";

            string envPythonLib = envPythonHome + @"\Lib\";
            string envPythonDll = envPythonHome + @"\DDLs\";

            string envDepPath = @"E:\Mark Diedericks\Documents\Visual Studio 2017\Projects\Excel Macros\Dependencies\";
            string envRunPath = @"E:\Mark Diedericks\Documents\Visual Studio 2017\Projects\Excel Macros\Excel Macros RIBBON\bin\x64\Debug\";

            string envPyNetPath = @"E:\Mark Diedericks\Documents\Visual Studio 2017\Projects\Excel Macros\Dependencies\Python .NET\pythonnet-2.3.0\";

            Environment.SetEnvironmentVariable("PYTHONHOME", envPythonHome, EnvironmentVariableTarget.Process);
            Environment.SetEnvironmentVariable("PATH", envPythonHome + ";" + Environment.GetEnvironmentVariable("PATH", EnvironmentVariableTarget.Machine), EnvironmentVariableTarget.Process);
            Environment.SetEnvironmentVariable("PYTHONPATH", envPythonLib + ';' + envPythonDll + ';' + envDepPath + ';' + envRunPath + ';' + envPyNetPath, EnvironmentVariableTarget.Process);

            PythonEngine.PythonHome = envPythonHome;
            PythonEngine.ProgramName = "Excel Macros Plus";

            PythonEngine.Initialize();
            IntPtr bat = PythonEngine.BeginAllowThreads();

            m_IsExecuting = false;
            m_BackgroundWorker = new BackgroundWorker();

            //Reset IO streams of ScriptEngine if they're changed
            Main.GetInstance().OnIOChanged += () =>
            {
                using (Py.GIL())
                {
                    dynamic sys = Py.Import("sys");
                    sys.stdout = new PyOutput(Main.GetEngineIOManager().GetOutput());
                    sys.stderr = new PyOutput(Main.GetEngineIOManager().GetError());
                }
            };

            //End running tasks if program is exiting
            Main.GetInstance().OnDestroyed += delegate () 
            {
                using (Py.GIL())
                {
                    PythonEngine.EndAllowThreads(bat);
                    PythonEngine.Shutdown();
                }

                if (m_BackgroundWorker != null)
                    m_BackgroundWorker.CancelAsync();
            };
        }

        #endregion

        #region Execution

        /// <summary>
        /// Determines how to execute a macro
        /// </summary>
        /// <param name="source">Source code (python)</param>
        /// <param name="OnCompletedAction">The action to be called once the code has been executed</param>
        /// <param name="async">If the code should be run asynchronously or not (synchronous)</param>
        /// <returns></returns>
        public bool ExecuteMacro(string source, Action OnCompletedAction, bool async)
        {
            if(m_IsExecuting)
                return false;

            if (async)
                ExecuteSourceAsynchronous(source, OnCompletedAction);
            else
                ExecuteSourceSynchronous(source, OnCompletedAction);

            return true;
        }

        /// <summary>
        /// End currently active asynchronous execution, if any
        /// </summary>
        public void TerminateExecution()
        {
            if (m_BackgroundWorker != null)
                m_BackgroundWorker.CancelAsync();

            m_IsExecuting = false;
        }

        /// <summary>
        /// Execute code asynchronously
        /// </summary>
        /// <param name="source">Source code (python)</param>
        /// <param name="OnCompletedAction">The action to be called once the code has been executed</param>
        private void ExecuteSourceAsynchronous(string source, Action OnCompletedAction)
        {
            m_IsExecuting = true;
            m_BackgroundWorker = new BackgroundWorker();
            m_BackgroundWorker.WorkerSupportsCancellation = true;
            int profileID = -1;

            m_BackgroundWorker.RunWorkerCompleted += (s, args) =>
            {
                if (Main.GetEngineIOManager() != null)
                {
                    Main.GetEngineIOManager().GetOutput().WriteLine("Asynchronous Execution Completed. Runtime of {0:N2}s", Utilities.GetTimeIntervalSeconds(profileID));
                    Main.GetEngineIOManager().GetOutput().Flush();
                }

                Utilities.EndProfileSession(profileID);

                m_IsExecuting = false;
                OnCompletedAction?.Invoke();
            };

            m_BackgroundWorker.DoWork += (s, args) =>
            {
                profileID = Utilities.BeginProfileSession();
                ExecuteSource(source);
            };

            m_BackgroundWorker.RunWorkerAsync();
        }

        /// <summary>
        /// Execute code synchronously
        /// </summary>
        /// <param name="source">Source code (python)</param>
        /// <param name="OnCompletedAction">The action to be called once the code has been executed</param>
        private void ExecuteSourceSynchronous(string source, Action OnCompletedAction)
        {
            Main.GetApplicationDispatcher().BeginInvoke(DispatcherPriority.Normal, new Action(() =>
            {
                int profileID = -1;
                profileID = Utilities.BeginProfileSession();

                m_IsExecuting = true;
                ExecuteSource(source);

                if (Main.GetEngineIOManager() != null)
                {
                    Main.GetEngineIOManager().GetOutput().WriteLine("Synchronous Execution Completed. Runtime of {0:N2}s", Utilities.GetTimeIntervalSeconds(profileID));
                    Main.GetEngineIOManager().GetOutput().Flush();
                }

                Utilities.EndProfileSession(profileID);

                m_IsExecuting = false;
                OnCompletedAction?.Invoke();
            }));
        }

        /// <summary>
        /// Execute source code through IronPython Script Engine
        /// </summary>
        /// <param name="source">Source code (python)</param>
        private void ExecuteSource(string source)
        {
            if (Main.GetEngineIOManager() != null)
                Main.GetEngineIOManager().ClearAllStreams();

            try
            {
                using (Py.GIL())
                {
                    using (PyScope scope = Py.CreateScope())
                    {
                        scope.Set("Utils", Utilities.GetExcelUtilities().ToPython());
                        scope.Set("Application", Utilities.GetInstance().GetApplication().ToPython());
                        scope.Set("ActiveWorkbook", Utilities.GetInstance().GetActiveWorkbook().ToPython());
                        scope.Set("ActiveWorksheet", Utilities.GetInstance().GetActiveWorksheet().ToPython());
                        scope.Set("MissingType", Type.Missing.ToPython());

                        scope.Exec(source);
                    }
                }
            }
            catch(ThreadAbortException tae)
            {
                System.Diagnostics.Debug.WriteLine("Execution Error: " + tae.Message);

                if (Main.GetEngineIOManager() != null)
                { 
                    Main.GetEngineIOManager().GetOutput().WriteLine("Thread Exited With Exception State {0}", tae.ExceptionState);
                    Main.GetEngineIOManager().GetOutput().Flush();
                }
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine("Execution Error: " + e.Message);

                if (Main.GetEngineIOManager() != null)
                { 
                    Main.GetEngineIOManager().GetError().WriteLine("Execution Error: " + e.Message);
                    Main.GetEngineIOManager().GetOutput().Flush();
                }
            }
        }

        #endregion
    }
}
