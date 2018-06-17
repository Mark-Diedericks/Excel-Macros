/*
 * Mark Diedericks
 * 17/06/2015
 * Version 1.0.0
 * Tool window view model
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace Excel_Macros_UI.ViewModel
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
            IsClosed = false;
        }

        public void Close()
        {
            IsClosed = true;
        }

        #region IsClosed

        private bool m_IsClosed;
        public bool IsClosed
        {
            get
            {
                return m_IsClosed;
            }

            set
            {
                if (m_IsClosed != value)
                {
                    m_IsClosed = value;
                    OnPropertyChanged(nameof(IsClosed));
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

        #region Title

        private string m_Title;
        public string Title
        {
            get
            {
                return m_Title;
            }

            set
            {
                if (m_Title != value)
                {
                    m_Title = value;
                    OnPropertyChanged(nameof(Title));
                }
            }
        }

        #endregion

    }
}
