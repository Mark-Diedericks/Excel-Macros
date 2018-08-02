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

namespace Excel_Macros_INTEROP.Engine
{
    public class ExecutionEngine
    {
        #region Static Initializaton

        /// <summary>
        /// Intialize static instances of the ExecutionEngine, one for 
        /// Debug execution and one for Release execution. No difference, 
        /// simply convention
        /// </summary>
        public static void Initialize()
        {
            Dictionary<string, object> debugArgs = new Dictionary<string, object>();
            Dictionary<string, object> releaseArgs = new Dictionary<string, object>();

            debugArgs["Debug"] = true;
            releaseArgs["Debug"] = false;

            s_DebugEngine = new ExecutionEngine(debugArgs);
            s_ReleaseEngine = new ExecutionEngine(releaseArgs);
        }

        private static ExecutionEngine s_DebugEngine;
        private static ExecutionEngine s_ReleaseEngine;

        /// <summary>
        /// Get instance of the Debug Execution Engine
        /// </summary>
        /// <returns>Debug Execution Engine</returns>
        public static ExecutionEngine GetDebugEngine()
        {
            if (s_DebugEngine == null)
                Initialize();

            return s_DebugEngine;
        }

        /// <summary>
        /// Get Instance of Release Execution Engine
        /// </summary>
        /// <returns>Release Execution Instance</returns>
        public static ExecutionEngine GetReleaseEngine()
        {
            if (s_ReleaseEngine == null)
                Initialize();

            return s_ReleaseEngine;
        }

        #endregion

        #region Instanced Initializaton

        private ScriptEngine m_ScriptEngine;
        private ScriptScope m_ScriptScope;

        private BackgroundWorker m_BackgroundWorker;
        private bool m_IsExecuting;
        
        /// <summary>
        /// Private Initialization of Exeuction Engine instance.
        /// </summary>
        /// <param name="args">Parameters to be used by IronPython Script Engine</param>
        private ExecutionEngine(Dictionary<string, object> args)
        {
            m_ScriptEngine = IronPython.Hosting.Python.CreateEngine(args);
            m_ScriptScope = m_ScriptEngine.CreateScope();

            m_IsExecuting = false;
            m_BackgroundWorker = new BackgroundWorker();

            //Reset IO streams of ScriptEngine if they're changed
            Main.GetInstance().OnIOChanged += () =>
            {
                m_ScriptEngine.Runtime.IO.RedirectToConsole();
                Console.SetOut(Main.GetEngineIOManager().GetOutput());
                Console.SetError(Main.GetEngineIOManager().GetError());
            };

            //End running tasks if program is exiting
            Main.GetInstance().OnDestroyed += delegate () 
            {
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
            object temp;
            if(!m_ScriptScope.TryGetVariable("Utils", out temp))
            {
                m_ScriptScope.SetVariable("Utils", Utilities.GetExcelUtilities());
                m_ScriptScope.SetVariable("Application", Utilities.GetInstance().GetApplication());
                m_ScriptScope.SetVariable("ActiveWorkbook", Utilities.GetInstance().GetActiveWorkbook());
                m_ScriptScope.SetVariable("ActiveWorksheet", Utilities.GetInstance().GetActiveWorksheet());
                m_ScriptScope.SetVariable("MissingType", Type.Missing);
            }

            if (Main.GetEngineIOManager() != null)
                Main.GetEngineIOManager().ClearAllStreams();

            try
            {
                m_ScriptEngine.Execute(source, m_ScriptScope);
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
