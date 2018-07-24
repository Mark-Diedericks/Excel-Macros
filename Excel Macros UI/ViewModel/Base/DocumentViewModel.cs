/*
 * Mark Diedericks
 * 19/07/2018
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
using System.Windows.Controls;
using System.Windows.Input;

namespace Excel_Macros_UI.ViewModel.Base
{
    public class DocumentViewModel : ViewModel
    {

        #region Close Command

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

        public void Close()
        {
            IsClosed = true;
        }

        #endregion

        #region Save Command

        private ICommand m_SaveCommand;
        public ICommand SaveCommand
        {
            get
            {
                if (m_SaveCommand == null)
                    m_SaveCommand = new RelayCommand((OnComplete) => Save((Action)OnComplete));
                return m_SaveCommand;
            }
        }

        public virtual void Save(Action OnComplete) { }

        #endregion

        #region Start Command

        private ICommand m_StartCommand;
        public ICommand StartCommand
        {
            get
            {
                if (m_StartCommand == null)
                    m_StartCommand = new RelayCommand((OnComplete) => Start((Action)OnComplete));
                return m_StartCommand;
            }
        }

        public virtual void Start(Action OnComplete) { }

        #endregion

        #region Stop Command

        private ICommand m_StopCommand;
        public ICommand StopCommand
        {
            get
            {
                if (m_StopCommand == null)
                    m_StopCommand = new RelayCommand((OnComplete) => Stop((Action)OnComplete));
                return m_StopCommand;
            }
        }

        public virtual void Stop(Action OnComplete) { }

        #endregion

        #region Undo Command

        private ICommand m_UndoCommand;
        public ICommand UndoCommand
        {
            get
            {
                return m_UndoCommand;
            }
            set
            {
                m_UndoCommand = value;
            }
        }

        #endregion

        #region Redo Command

        private ICommand m_RedoCommand;
        public ICommand RedoCommand
        {
            get
            {
                return m_RedoCommand;
            }
            set
            {
                m_RedoCommand = value;
            }
        }

        #endregion

        public DocumentViewModel()
        {
            CanClose = true;
            CanFloat = false;
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

        #region ToolTip

        public string ToolTip
        {
            get
            {
                return Model.ToolTip;
            }

            set
            {
                if (Model.ToolTip != value)
                {
                    Model.ToolTip = value;
                    OnPropertyChanged(nameof(ToolTip));
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
