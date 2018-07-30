using Excel_Macros_INTEROP;
using Excel_Macros_INTEROP.Macros;
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
        public TextualEditorModel(Guid id)
        {
            if(id != Guid.Empty)
            {
                IMacro macro = Main.GetMacro(id);

                if (macro != null)
                {
                    Title = Main.GetMacro(id).GetName();
                    ToolTip = Main.GetMacro(id).GetRelativePath();
                    ContentId = Main.GetMacro(id).GetRelativePath();
                    Macro = id;
                    IsClosed = false;
                    Source = new TextDocument(Main.GetMacro(id).GetSource());
                    IsSaved = true;
                    return;
                }
            }

            Source = new TextDocument();
            IsSaved = false;
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
                    m_Source.TextChanged += (s, e) => { IsSaved = false; };
                    OnPropertyChanged(nameof(Source));
                    IsSaved = false;
                }
            }
        }

        #endregion
    }
}
