using Excel_Macros_INTEROP;
using Excel_Macros_INTEROP.Macros;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_UI.Model.Base
{
    public class DocumentModel : Model
    {
        public DocumentModel()
        {
            IsClosed = false;
            Title = "";
            ToolTip = "";
            ContentId = "";
            Macro = Guid.Empty;
        }

        public static DocumentModel Create(Guid id)
        {
            MacroDeclaration declaration = Main.GetDeclaration(id);

            if (declaration == null)
                return null;

            switch(declaration.type)
            {
                case MacroType.PYTHON:
                    return new TextualEditorModel(id);
                case MacroType.BLOCKLY:
                    return new VisualEditorModel(id);
            }

            return null;
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

        #region ToolTip

        private string m_ToolTip;
        public string ToolTip
        {
            get
            {
                return m_ToolTip;
            }

            set
            {
                if (m_ToolTip != value)
                {
                    m_ToolTip = value;
                    OnPropertyChanged(nameof(ToolTip));
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

        #region Macro

        private Guid m_Macro;
        public Guid Macro
        {
            get
            {
                return m_Macro;
            }

            set
            {
                if (m_Macro != value)
                {
                    m_Macro = value;
                    OnPropertyChanged(nameof(Macro));
                }
            }
        }

        #endregion

    }
}
