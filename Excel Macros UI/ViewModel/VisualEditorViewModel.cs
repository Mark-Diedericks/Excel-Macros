/*
 * Mark Diedericks
 * 02/08/2018
 * Version 1.0.3
 * Visual editor view model
 */

using Excel_Macros_INTEROP;
using Excel_Macros_INTEROP.Macros;
using Excel_Macros_UI.Model;
using Excel_Macros_UI.Routing;
using Excel_Macros_UI.ViewModel.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace Excel_Macros_UI.ViewModel
{
    public class VisualEditorViewModel : DocumentViewModel
    {
        /// <summary>
        /// Instantiation of VisualEditorViewModel
        /// </summary>
        public VisualEditorViewModel()
        {
            Model = new VisualEditorModel(Guid.Empty);
            IsSaved = true;
            Title = "Visual Editor";
            ToolTip = "Prototyping Scratchpad";
            ContentId = "Visual Editor";
            IsClosed = false;
            CanClose = false;
            Source = String.Empty;
        }

        /// <summary>
        /// Saves the macro associated with the document
        /// </summary>
        /// <param name="OnComplete">Action to be fired on the tasks completetion</param>
        public override void Save(Action OnComplete)
        {
            base.Stop(OnComplete);
        }

        /// <summary>
        /// Executes the macro associated with the document
        /// </summary>
        /// <param name="OnComplete">Action to be fired on the tasks completetion</param>
        public override void Start(Action OnComplete)
        {
            Excel_Macros_INTEROP.Engine.ExecutionEngine.GetEngine().ExecuteMacro(GetPythonCode(), OnComplete, MainWindowViewModel.GetInstance().AsyncExecution);
            base.Start(null);
        }

        /// <summary>
        /// Terminates the execution of the macro associated with the document
        /// </summary>
        /// <param name="OnComplete">Action to be fired on the tasks completetion</param>
        public override void Stop(Action OnComplete)
        {
            Excel_Macros_INTEROP.Engine.ExecutionEngine.GetEngine().TerminateExecution();
            base.Stop(null);
        }

        public delegate string InvokeEngineEvent();
        public event InvokeEngineEvent InvokeEngine;

        /// <summary>
        /// Gets the python code of the visual program
        /// </summary>
        /// <returns>Source code (python)</returns>
        public string GetPythonCode()
        {
            return InvokeEngine?.Invoke();
        }

        #region Model

        public new VisualEditorModel Model
        {
            get
            {
                return (VisualEditorModel)base.Model;
            }

            set
            {
                if (((VisualEditorModel)base.Model) != value)
                {
                    base.Model = value;
                    OnPropertyChanged(nameof(Model));
                }
            }
        }

        #endregion

        #region Source

        public string Source
        {
            get
            {
                return Model.Source;
            }

            set
            {
                if (Model.Source != value)
                {
                    Model.Source = value;
                    OnPropertyChanged(nameof(Source));
                }
            }
        }

        #endregion

        #region WebBrowserVisibility

        public Visibility m_WebBrowserVisibility;
        public Visibility WebBrowserVisibility
        {
            get
            {
                return m_WebBrowserVisibility;
            }

            set
            {
                if (m_WebBrowserVisibility != value)
                {
                    m_WebBrowserVisibility = value;
                    OnPropertyChanged(nameof(WebBrowserVisibility));
                }
            }
        }

        #endregion

    }
}
