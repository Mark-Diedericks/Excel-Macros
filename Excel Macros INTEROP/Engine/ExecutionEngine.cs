/*
 * Mark Diedericks
 * 22/07/2018
 * Version 1.0.8
 * Manages execution of users' code
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

        public static ExecutionEngine GetDebugEngine()
        {
            if (s_DebugEngine == null)
                Initialize();

            return s_DebugEngine;
        }

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

        public object DispatcherPriorty { get; private set; }

        private ExecutionEngine(Dictionary<string, object> args)
        {
            m_ScriptEngine = IronPython.Hosting.Python.CreateEngine(args);
            m_ScriptScope = m_ScriptEngine.CreateScope();

            m_IsExecuting = false;
            m_BackgroundWorker = new BackgroundWorker();

            Main.GetInstance().OnIOChanged += () =>
            {
                m_ScriptEngine.Runtime.IO.RedirectToConsole();
                Console.SetOut(Main.GetEngineIOManager().GetOutput());
                Console.SetError(Main.GetEngineIOManager().GetError());
            };

            Main.GetInstance().OnDestroyed += delegate () 
            {
                if (m_BackgroundWorker != null)
                    m_BackgroundWorker.CancelAsync();
            };
        }

        #endregion

        #region Execution

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

        public void TerminateExecution()
        {
            if (m_BackgroundWorker != null)
                m_BackgroundWorker.CancelAsync();

            m_IsExecuting = false;
        }

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
                    Main.GetEngineIOManager().GetOutput().WriteLine("Syncrhonous Execution Completed. Runtime of {0:N2}s", Utilities.GetTimeIntervalSeconds(profileID));
                    Main.GetEngineIOManager().GetOutput().Flush();
                }

                Utilities.EndProfileSession(profileID);

                m_IsExecuting = false;
                OnCompletedAction?.Invoke();
            }));
        }

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
