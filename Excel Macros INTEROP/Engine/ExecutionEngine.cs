/*
 * Mark Diedericks
 * 04/07/2018
 * Version 1.0.4
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

        private ExecutionEngine(Dictionary<string, object> args)
        {
            m_ScriptEngine = IronPython.Hosting.Python.CreateEngine(args);
            m_ScriptScope = m_ScriptEngine.CreateScope();

            m_IsExecuting = false;
            m_BackgroundWorker = new BackgroundWorker();

            StreamManager.ClearAllStreams();
            m_ScriptEngine.Runtime.IO.SetInput(StreamManager.GetInputStream(), Encoding.UTF8);
            m_ScriptEngine.Runtime.IO.SetOutput(StreamManager.GetOutputStream(), Encoding.UTF8);
            m_ScriptEngine.Runtime.IO.SetErrorOutput(StreamManager.GetErrorStream(), Encoding.UTF8);

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
            m_BackgroundWorker.WorkerSupportsCancellation = true;

            m_BackgroundWorker.RunWorkerCompleted += (s, args) => 
            {
                OnCompletedAction?.Invoke();
                m_IsExecuting = false;
            };

            m_BackgroundWorker.DoWork += (s, args) =>
            {
                ExecuteSource(source);
            };

            m_BackgroundWorker.RunWorkerAsync();
        }

        private void ExecuteSourceSynchronous(string source, Action OnCompletedAction)
        {
            m_IsExecuting = true;
            ExecuteSource(source);
            OnCompletedAction?.Invoke();
            m_IsExecuting = false;
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

            StreamManager.ClearAllStreams();

            try
            {
                m_ScriptEngine.Execute(source, m_ScriptScope);
            }
            catch(ThreadAbortException tae)
            {
                System.Diagnostics.Debug.WriteLine("Execution Error: " + tae.Message);
                StreamManager.GetOutputWriter().WriteLine("Thread Exited With Exception State {0}", tae.ExceptionState);
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.WriteLine("Execution Error: " + e.Message);
                StreamManager.GetErrorWriter().WriteLine("Execution Error: " + e.Message);
            }
        }

        #endregion
    }
}
