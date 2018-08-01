/*
 * Mark Diedericks
 * 30/07/2018
 * Version 1.0.1
 * Visual Macro data structure
 */

using Excel_Macros_INTEROP.Engine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_INTEROP.Macros
{
    public class VisualMacro : IMacro
    {
        private Guid m_ID;
        private string m_Source;
        private string m_Python;

        public VisualMacro(string source)
        {
            m_Source = source;
        }

        public void CreateBlankMacro()
        {
            m_Source = "<xml>\n</xml>";
        }

        public Guid GetID()
        {
            return m_ID;
        }

        public void SetID(Guid id)
        {
            m_ID = id;
        }

        public void SetSource(string source)
        {
            m_Source = source;
            EventManager.ConvertToPython(m_ID, (s) => m_Python = s);
        }

        public string GetSource()
        {
            return m_Source;
        }

        public void Rename(string name)
        {
            FileManager.RenameMacro(m_ID, name);

            if (Main.IsRibbonMacro(m_ID))
                Main.RenameRibbonMacro(m_ID);
        }

        public string GetName()
        {
            return Main.GetDeclaration(m_ID).name;
        }

        public string GetRelativePath()
        {
            return Main.GetDeclaration(m_ID).relativepath;
        }

        public void Save()
        {
            FileManager.SaveMacro(m_ID, m_Source);
        }

        public void Export()
        {
            FileManager.ExportMacro(m_ID, m_Source);
        }

        public void Delete(Action<bool> OnReturn)
        {
            FileManager.DeleteMacro(m_ID, OnReturn);

            if (Main.IsRibbonMacro(m_ID))
                Main.RemoveRibbonMacro(m_ID);
        }

        public void ExecuteDebug(Action OnCompletedAction, bool async)
            {
                if (!String.IsNullOrEmpty(m_Python))
                    ExecutionEngine.GetDebugEngine().ExecuteMacro(m_Python, OnCompletedAction, async);
            }

        public void ExecuteRelease(Action OnCompletedAction, bool async)
        {
            if(!String.IsNullOrEmpty(m_Python))
                ExecutionEngine.GetReleaseEngine().ExecuteMacro(m_Python, OnCompletedAction, async);
        }
    }
}
