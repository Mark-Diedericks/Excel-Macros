/*
 * Mark Diedericks
 * 08/06/2018
 * Version 1.0.0
 * Visual Macro data structure
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_INTEROP.Macros
{
    public class VisualMacro : IMacro
    {
        public VisualMacro(string source)
        {

        }

        public void CreateBlankMacro()
        {
            throw new NotImplementedException();
        }

        public void Delete(Action<bool> OnReturn)
        {
            throw new NotImplementedException();
        }

        public void ExecuteDebug(Action OnCompletedAction, bool async)
        {
            throw new NotImplementedException();
        }

        public void ExecuteRelease(Action OnCompletedAction, bool async)
        {
            throw new NotImplementedException();
        }

        public void Export()
        {
            throw new NotImplementedException();
        }

        public Guid GetID()
        {
            throw new NotImplementedException();
        }

        public string GetName()
        {
            throw new NotImplementedException();
        }

        public string GetRelativePath()
        {
            throw new NotImplementedException();
        }

        public string GetSource()
        {
            throw new NotImplementedException();
        }

        public void Rename(string name)
        {
            throw new NotImplementedException();
        }

        public void Save()
        {
            throw new NotImplementedException();
        }

        public void SetID(Guid id)
        {
            throw new NotImplementedException();
        }

        public void SetSource(string source)
        {
            throw new NotImplementedException();
        }
    }
}
