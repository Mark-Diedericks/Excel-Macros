/*
 * Mark Diedericks
 * 17/06/2015
 * Version 1.0.0
 * Textual editor view model
 */

using Excel_Macros_UI.Model;
using Excel_Macros_UI.ViewModel.Base;
using ICSharpCode.AvalonEdit.Document;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_UI.ViewModel
{
    public class TextualEditorViewModel : DocumentViewModel
    {
        public TextualEditorViewModel()
        {
            Model = new TextualEditorModel();
            Source = new TextDocument();
            Macro = Guid.Empty;
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

        #region Macro

        public Guid Macro
        {
            get
            {
                return Model.Macro;
            }

            set
            {
                if (Model.Macro != value)
                {
                    Model.Macro = value;
                    OnPropertyChanged(nameof(Macro));
                }
            }
        }

        #endregion
    }
}
