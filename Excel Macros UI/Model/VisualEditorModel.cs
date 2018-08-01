/*
 * Mark Diedericks
 * 30/07/2018
 * Version 1.0.2
 * Visual editor model, data handling
 */

using Excel_Macros_INTEROP;
using Excel_Macros_INTEROP.Macros;
using Excel_Macros_UI.Model.Base;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_UI.Model
{
    public class VisualEditorModel : DocumentModel
    {

        public VisualEditorModel(Guid id)
        {
            if (id != Guid.Empty)
            {
                IMacro macro = Main.GetMacro(id);

                if (macro != null)
                {
                    Title = Main.GetMacro(id).GetName();
                    ToolTip = Main.GetMacro(id).GetRelativePath();
                    ContentId = Main.GetMacro(id).GetRelativePath();
                    Macro = id;
                    IsClosed = false;
                    Source = Main.GetMacro(id).GetSource();
                    IsSaved = true;
                    return;
                }
            }

            Source = String.Empty;
            IsSaved = false;
        }

        #region Source

        private string m_Source;
        public string Source
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
                    IsSaved = false;
                    OnPropertyChanged(nameof(Source));
                }
            }
        }

        #endregion

    }
}
