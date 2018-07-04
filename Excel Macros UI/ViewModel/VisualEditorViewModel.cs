/*
 * Mark Diedericks
 * 17/06/2015
 * Version 1.0.0
 * Visual editor view model
 */

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
            Model = new VisualEditorModel();
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

    }
}
