/*
 * Mark Diedericks
 * 17/06/2018
 * Version 1.0.0
 * Visual editor view model
 */

using Excel_Macros_INTEROP;
using Excel_Macros_INTEROP.Macros;
using Excel_Macros_UI.Model;
using Excel_Macros_UI.ViewModel.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_UI.ViewModel
{
    public class VisualEditorViewModel : DocumentViewModel
    {

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

        public override void Save(Action OnComplete)
        {
            //Main.GetMacro(Macro).SetSource(Source);
            //Main.GetMacro(Macro).Save();
            base.Stop(OnComplete);
        }

        public override void Start(Action OnComplete)
        {
            /*IMacro macro = Main.GetMacro(Macro);
            MacroDeclaration declaration = Main.GetDeclaration(Macro);

            if (macro == null || declaration == null || declaration.type != MacroType.BLOCKLY)
            {
                OnComplete?.Invoke();
                return;
            }

            macro.SetSource(Source);
            macro.ExecuteDebug(OnComplete, MainWindowViewModel.GetInstance().AsyncExecution);*/

            Excel_Macros_INTEROP.Engine.ExecutionEngine.GetDebugEngine().ExecuteMacro(GetPythonCode(), OnComplete, MainWindowViewModel.GetInstance().AsyncExecution);
            base.Start(null);
        }

        public override void Stop(Action OnComplete)
        {
            Excel_Macros_INTEROP.Engine.ExecutionEngine.GetDebugEngine().TerminateExecution();
            base.Stop(null);
        }

        public delegate string InvokeEngineEvent();
        public event InvokeEngineEvent InvokeEngine;

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

    }
}
