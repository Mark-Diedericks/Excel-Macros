/*
 * Mark Diedericks
 * 22/07/2018
 * Version 1.0.3
 * Abstraction layer; base macro interface
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_INTEROP.Macros
{
    public interface IMacro
    {
        void CreateBlankMacro();

        void SetID(Guid id);
        Guid GetID();

        void SetSource(string source);
        string GetSource();

        void ExecuteDebug(Action OnCompletedAction, bool async);
        void ExecuteRelease(Action OnCompletedAction, bool async);

        void Save();
        void Export();

        void Delete(Action<bool> OnReturn);

        void Rename(string name);
        string GetName();

        string GetRelativePath();
    }
}
