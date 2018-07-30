/*
 * Mark Diedericks
 * 17/06/2018
 * Version 1.0.0
 * Tool window view model
 */

using Excel_Macros_UI.Model.Base;
using Excel_Macros_UI.Routing;
using Excel_Macros_UI.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Input;

namespace Excel_Macros_UI.ViewModel.Base
{
    public class ToolViewModel : ViewModel
    {
        public ToolViewModel()
        {

        }

        #region Model

        private ToolModel m_Model;
        public ToolModel Model
        {
            get
            {
                return m_Model;
            }

            set
            {
                if (m_Model != value)
                {
                    m_Model = value;
                    OnPropertyChanged(nameof(Model));
                }
            }
        }

        #endregion

        #region IsVisible

        public bool IsVisible
        {
            get
            {
                return Model.IsVisible;
            }

            set
            {
                if (Model.IsVisible != value)
                {
                    Model.IsVisible = value;
                    OnPropertyChanged(nameof(IsVisible));
                }
            }
        }

        #endregion

        #region Title

        public string Title
        {
            get
            {
                return Model.Title;
            }

            set
            {
                if (Model.Title != value)
                {
                    Model.Title = value;
                    OnPropertyChanged(nameof(Title));
                }
            }
        }

        #endregion

        #region ContentId
        
        public string ContentId
        {
            get
            {
                return Model.ContentId;
            }

            set
            {
                if (Model.ContentId != value)
                {
                    Model.ContentId = value;
                    OnPropertyChanged(nameof(ContentId));
                }
            }
        }

        #endregion

    }
}
