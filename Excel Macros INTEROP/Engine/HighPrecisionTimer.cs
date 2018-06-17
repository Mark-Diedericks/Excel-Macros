﻿/*
 * Mark Diedericks
 * 09/06/2018
 * Version 1.0.0
 * A high precision timing class
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Excel_Macros_INTEROP.Engine
{
    internal class HighPrecisionTimer
    {
        //
        // From:     CodeProject
        // Author:   Daniel Strigl
        // Link:     https://www.codeproject.com/Articles/2635/High-Performance-Timer-in-C
        // Lisence:  https://www.codeproject.com/info/cpol10.aspx
        //
        [DllImport("Kernel32.dll")]
        private static extern bool QueryPerformanceCounter(out long lpPerformanceCount);

        [DllImport("Kernel32.dll")]
        private static extern bool QueryPerformanceFrequency(out long lpFrequency);

        private long startTime, stopTime;
        private long freq;

        // Constructor
        public HighPrecisionTimer()
        {
            startTime = 0;
            stopTime = 0;

            if (QueryPerformanceFrequency(out freq) == false)
            {
                // high-performance counter not supported
                throw new Win32Exception();
            }
        }

        // Start the timer
        public void Start()
        {
            // lets do the waiting threads there work
            Thread.Sleep(0);

            QueryPerformanceCounter(out startTime);
        }

        // Stop the timer
        public void Stop()
        {
            QueryPerformanceCounter(out stopTime);
        }

        // Returns the duration of the timer (in ms)
        public double Duration
        {
            get
            {
                return ((double)(stopTime - startTime) / (double)freq) * 1000.0D;
            }
        }
    }
}
