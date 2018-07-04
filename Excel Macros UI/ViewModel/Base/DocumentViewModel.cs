/*
 * Mark Diedericks
 * 17/06/2015
 * Version 1.0.0
 * Document window view model
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
    public class DocumentViewModel : ViewModel
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

        public DocumentViewModel()
        {
            CanClose = true;
            CanFloat = false;
        }

        public void Close()
        {
            IsClosed = true;
        }

        #region Model

        private DocumentModel m_Model;
        public DocumentModel Model
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

        #region CanFloat

        private bool m_CanFloat;
        public bool CanFloat
        {
            get
            {
                return m_CanFloat;
            }

            set
            {
                if (m_CanFloat != value)
                {
                    m_CanFloat = value;
                    OnPropertyChanged(nameof(CanFloat));
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
                if(Model.Title != value)
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
