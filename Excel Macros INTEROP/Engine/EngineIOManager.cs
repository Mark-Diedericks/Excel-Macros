/*
 * Mark Diedericks
 * 22/07/2018
 * Version 1.0.0
 * Manages execution engines' IO
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_INTEROP.Engine
{
    public class EngineIOManager
    {
        private TextWriter Output;
        private TextWriter Error;

        public EngineIOManager(TextWriter output, TextWriter error)
        {
            Output = output;
            Error = error;
        }

        public void ClearAllStreams()
        {
            Output.Flush();
            Error.Flush();

            EventManager.ClearAllIO();
        }

        public TextWriter GetOutput()
        {
            return Output;
        }

        public TextWriter GetError()
        {
            return Error;
        }

    }
}
