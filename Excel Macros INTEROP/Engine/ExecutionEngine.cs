/*
 * Mark Diedericks
 * 09/06/2015
 * Version 1.0.2
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

namespace Excel_Macros_INTEROP.Engine
{
    public class ExecutionEngine
    {
        #region Static Initializaton

        public static void Initialize()
        {
            Main.GetInteropDispatcher().Invoke(() =>
            {
                Dictionary<string, object> debugArgs = new Dictionary<string, object>();
                Dictionary<string, object> releaseArgs = new Dictionary<string, object>();

                debugArgs["Debug"] = true;
                releaseArgs["Debug"] = false;

                s_DebugEngine = new ExecutionEngine(debugArgs);
                s_ReleaseEngine = new ExecutionEngine(releaseArgs);
            });
        }

        private static ExecutionEngine s_DebugEngine;
        private static ExecutionEngine s_ReleaseEngine;

        public static ExecutionEngine GetDebugEngine()
        {
            return Main.GetInteropDispatcher().Invoke(() =>
            {
                if (s_DebugEngine == null)
                    Initialize();

                return s_DebugEngine;
            });
        }

        public static ExecutionEngine GetReleaseEngine()
        {
            return Main.GetInteropDispatcher().Invoke(() =>
            {
                if (s_ReleaseEngine == null)
                    Initialize();

                return s_ReleaseEngine;
            });
        }

        #endregion

        #region Instanced Initializaton

        private ScriptEngine m_ScriptEngine;
        private ScriptScope m_ScriptScope;

        private Thread m_ExecutionThread;

        private ExecutionEngine(Dictionary<string, object> args)
        {
            m_ScriptEngine = IronPython.Hosting.Python.CreateEngine(args);
            m_ScriptScope = m_ScriptEngine.CreateScope();

            m_ExecutionThread = null;

            StreamManager.ClearAllStreams();
            m_ScriptEngine.Runtime.IO.SetInput(StreamManager.GetInputStream(), Encoding.UTF8);
            m_ScriptEngine.Runtime.IO.SetOutput(StreamManager.GetOutputStream(), Encoding.UTF8);
            m_ScriptEngine.Runtime.IO.SetErrorOutput(StreamManager.GetErrorStream(), Encoding.UTF8);
        }

        #endregion

        #region Execution

        public bool ExecuteMacro(string source, Action OnCompletedAction, bool async)
        {
            if (m_ExecutionThread != null)
                return false;

            if (async)
                Main.GetExcelDispatcher().InvokeAsync(() => ExecuteSourceAsynchronous(source, OnCompletedAction));
            else
                Main.GetExcelDispatcher().InvokeAsync(() => ExecuteSource(source, OnCompletedAction));

            return true;
        }

        public void TerminateExecution()
        {
            if (m_ExecutionThread != null)
                Main.GetExcelDispatcher().BeginInvoke(DispatcherPriority.Send, new Action(() => m_ExecutionThread.Abort()));
        }

        private void ExecuteSourceAsynchronous(string source, Action OnCompletedAction)
        {
            m_ExecutionThread = new Thread((ThreadStart)delegate
            {
                ExecuteSource(source, OnCompletedAction);
                Main.GetExcelDispatcher().BeginInvoke(DispatcherPriority.Normal, new Action(() => m_ExecutionThread = null));
            });

            m_ExecutionThread.Start();
        }

        private void ExecuteSource(string source, Action OnCompletedAction)
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

            Main.GetInteropDispatcher().BeginInvoke(DispatcherPriority.Normal, new Action(() => OnCompletedAction?.Invoke()));
        }

        #endregion
    }
}
