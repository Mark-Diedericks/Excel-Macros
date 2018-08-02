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

        /// <summary>
        /// Initialize EngineIOManager
        /// </summary>
        /// <param name="output">Console Output TextWriter</param>
        /// <param name="error">Console Error TextWriter</param>
        public EngineIOManager(TextWriter output, TextWriter error)
        {
            Output = output;
            Error = error;
        }

        /// <summary>
        /// Clear all TextWriters of data and clear the UI of displayed text
        /// </summary>
        public void ClearAllStreams()
        {
            Output.Flush();
            Error.Flush();

            EventManager.ClearAllIO();
        }

        /// <summary>
        /// Get the TextWriter object for console outputs
        /// </summary>
        /// <returns>Output TextWriter</returns>
        public TextWriter GetOutput()
        {
            return Output;
        }

        /// <summary>
        /// Get the TextWriter object for console errors
        /// </summary>
        /// <returns>Error TextWriter</returns>
        public TextWriter GetError()
        {
            return Error;
        }

    }
}
