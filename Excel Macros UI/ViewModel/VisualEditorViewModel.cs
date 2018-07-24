/*
 * Mark Diedericks
 * 17/06/2018
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
            Model = new VisualEditorModel(Guid.Empty);
        }

        public override void Save(Action OnComplete)
        {
            throw new NotImplementedException();
            OnComplete?.Invoke();
        }

        public override void Start(Action OnComplete)
        {
            throw new NotImplementedException();
            OnComplete?.Invoke();
        }

        public override void Stop(Action OnComplete)
        {
            throw new NotImplementedException();
            OnComplete?.Invoke();
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
