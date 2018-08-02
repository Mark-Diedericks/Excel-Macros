/*
 * Mark Diedericks
 * 09/06/2018
 * Version 1.0.0
 * A data structure to store basic timing/profiling info
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Macros_INTEROP.Engine
{
    public struct ProfileInfo
    {
        public double Duration;
        public string Statement;
        public int LineIndex;

        /// <summary>
        /// Initialize a new instance of the data structure
        /// </summary>
        /// <param name="duration">Duration of the profile (ms)</param>
        /// <param name="statement">Code which was profiled (python source)</param>
        /// <param name="line">Line number at which the code started in the respective file</param>
        public ProfileInfo(double duration, string statement, int line)
        {
            Duration = duration;
            Statement = statement;
            LineIndex = line;
        }
    }
}
