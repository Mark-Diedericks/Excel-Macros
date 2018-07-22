/*
 * Mark Diedericks
 * 17/06/2018
 * Version 1.0.0
 * Tool window view model
 */

using Excel_Macros_UI.Model.Base;
using Excel_Macros_UI.Routing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace Excel_Macros_UI.ViewModel.Base
{
    public class ToolViewModel : ViewModel
    {
        private ICommand m_CloseCommand;
        public ICommand CloseCommand
        {
            get
            {
                if (m_CloseCommand == null)
                    m_CloseCommand = new RelayCommand(call => Close());
                return m_CloseCommand;
            }
        }

        public ToolViewModel()
        {
            CanClose = true;
        }

        public void Close()
        {
            IsClosed = true;
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

        #region CanClose

        private bool m_CanClose;
        public bool CanClose
        {
            get
            {
                return m_CanClose;
            }

            set
            {
                if (m_CanClose != value)
                {
                    m_CanClose = value;
                    OnPropertyChanged(nameof(CanClose));
                }
            }
        }

        #endregion

        #region IsClosed
        
        public bool IsClosed
        {
            get
            {
                return Model.IsClosed;
            }

            set
            {
                if (Model.IsClosed != value)
                {
                    Model.IsClosed = value;
                    OnPropertyChanged(nameof(IsClosed));
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
