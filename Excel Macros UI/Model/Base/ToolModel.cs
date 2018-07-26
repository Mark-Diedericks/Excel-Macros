using Excel_Macros_UI.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_UI.Model.Base
{
    public class ToolModel : Model
    {
        public ToolModel()
        {
            IsVisible = true;
            m_PreferredLocation = PaneLocation.Left;
            Title = "";
            ContentId = "";
        }

        #region IsVisible

        private bool m_IsVisible;
        public bool IsVisible
        {
            get
            {
                return m_IsVisible;
            }
            set
            {
                if (m_IsVisible != value)
                {
                    m_IsVisible = value;
                    OnPropertyChanged(nameof(IsVisible));
                }
            }
        }

        #endregion

        #region PreferredLocation

        private PaneLocation m_PreferredLocation;
        public PaneLocation PreferredLocation
        {
            get
            {
                return m_PreferredLocation;
            }
            set
            {
                if (m_PreferredLocation != value)
                {
                    m_PreferredLocation = value;
                    OnPropertyChanged(nameof(PreferredLocation));
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

        #region ContentId

        private string m_ContentId;
        public string ContentId
        {
            get
            {
                return m_ContentId;
            }

            set
            {
                if (m_ContentId != value)
                {
                    m_ContentId = value;
                    OnPropertyChanged(nameof(ContentId));
                }
            }
        }

        #endregion   

    }
}
