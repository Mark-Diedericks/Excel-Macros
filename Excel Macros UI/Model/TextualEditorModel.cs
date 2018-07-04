using Excel_Macros_UI.Model.Base;
using ICSharpCode.AvalonEdit.Document;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_UI.Model
{
    public class TextualEditorModel : DocumentModel
    {
        public TextualEditorModel()
        {
            Source = new TextDocument();
            Macro = Guid.Empty;
        }

        #region Source

        private TextDocument m_Source;
        public TextDocument Source
        {
            get
            {
                return m_Source;
            }

            set
            {
                if (m_Source != value)
                {
                    m_Source = value;
                    OnPropertyChanged(nameof(Source));
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
