/*
 * Mark Diedericks
 * 30/07/2018
 * Version 1.0.3
 * Textual editor view model
 */

using Excel_Macros_INTEROP;
using Excel_Macros_UI.Model;
using Excel_Macros_UI.Routing;
using Excel_Macros_UI.View;
using Excel_Macros_UI.ViewModel.Base;
using ICSharpCode.AvalonEdit.Document;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace Excel_Macros_UI.ViewModel
{
    public class TextualEditorViewModel : DocumentViewModel
    {
        public TextualEditorViewModel()
        {
            Model = new TextualEditorModel(Guid.Empty);
            IsSaved = true;
        }

        public override void Save(Action OnComplete)
        {
            Main.GetMacro(Macro).SetSource(Source.Text);
            Main.GetMacro(Macro).Save();
            base.Stop(OnComplete);
        }

        public override void Start(Action OnComplete)
        {
            Excel_Macros_INTEROP.Engine.ExecutionEngine.GetDebugEngine().ExecuteMacro(Source.Text, OnComplete, MainWindowViewModel.GetInstance().AsyncExecution);
            base.Stop(null);
        }

        public override void Stop(Action OnComplete)
        {
            Excel_Macros_INTEROP.Engine.ExecutionEngine.GetDebugEngine().TerminateExecution();
            base.Stop(null);
        }

        #region Model

        public new TextualEditorModel Model
        {
            get
            {
                return (TextualEditorModel)base.Model;
            }

            set
            {
                if (((TextualEditorModel)base.Model) != value)
                {
                    base.Model = value;
                    OnPropertyChanged(nameof(Model));
                }
            }
        }

        #endregion

        #region Source

        public TextDocument Source
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
