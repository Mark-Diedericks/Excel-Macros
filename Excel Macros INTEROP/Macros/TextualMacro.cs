﻿/*
 * Mark Diedericks
 * 09/06/2015
 * Version 1.0.0
 * Textual Macro data structure
 */

using Excel_Macros_INTEROP.Engine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_INTEROP.Macros
{
    public class TextualMacro : IMacro
    {
        private Guid m_ID;
        private string m_Source;

        public TextualMacro(string source)
        {
            m_Source = source;
            m_ID = Guid.Empty;
        }

        public void CreateBlankMacro()
        {
            m_Source = "import clr" + "\n\n\n";
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
        }

        public string GetSource()
        {
            return m_Source;
        }

        public void Rename(string name)
        {
            FileManager.RenameMacro(m_ID, name);
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
        }

        public void ExecuteDebug(Action OnCompletedAction, bool async)
        {
            ExecutionEngine.GetDebugEngine().ExecuteMacro(m_Source, OnCompletedAction, async);
        }

        public void ExecuteRelease(Action OnCompletedAction, bool async)
        {
            ExecutionEngine.GetReleaseEngine().ExecuteMacro(m_Source, OnCompletedAction, async);
        }
    }
}
