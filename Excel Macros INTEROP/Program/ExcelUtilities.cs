/*
 * Mark Diedericks
 * 22/07/2018
 * Version 1.0.2
 * Excel related utility functions for Users' use
 */

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_Macros_INTEROP
{
    public class ExcelUtilities
    {
        /// <summary>
        /// Get the Excel application
        /// </summary>
        /// <returns>Excel application instance</returns>
        public Excel.Application GetApplication()
        {
            return Main.GetApplication();
        }

        /// <summary>
        /// Get the active workbook
        /// </summary>
        /// <returns>Excel's ActiveWorkbook</returns>
        public Excel.Workbook GetActiveWorkbook()
        {
            return GetApplication().ActiveWorkbook;
        }

        /// <summary>
        /// Get the active worksheet
        /// </summary>
        /// <returns>Excel's ActiveWorksheet</returns>
        public Excel.Worksheet GetActiveWorksheet()
        {
            return (Excel.Worksheet)GetActiveWorkbook().ActiveSheet;
        }

        /// <summary>
        /// Get a range selection through the excel inputbox
        /// </summary>
        /// <param name="message">The message to be displayed</param>
        /// <returns>Selected Range</returns>
        public Excel.Range RequestRangeInput(string message)
        {
            object result = MessageManager.DisplayInputMessage(message, "Input Range", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
            Main.FireFocusEvent();
            return result as Excel.Range;
        }

        /// <summary>
        /// Get a boolean input through a message box
        /// </summary>
        /// <param name="message">The message to be displayed</param>
        /// <returns>Bool representation of Yes/No selection</returns>
        public bool RequestBooleanInput(string message)
        {
            bool result = MessageManager.DisplayYesNoMessage(message, "Boolean Input");
            Main.FireFocusEvent();
            return result;
        }

        /// <summary>
        /// Display a message in a message box
        /// </summary>
        /// <param name="message">The message to be displayed</param>
        /// <param name="caption">The header of the message</param>
        public void DisplayMessage(string message, string caption)
        {
            MessageManager.DisplayOkMessage(message, caption);
            Main.FireFocusEvent();
        }

        /// <summary>
        /// Sets FastWorkbook custom function, disable some Excel UI functions to improve performance.
        /// </summary>
        /// <param name="enable">Bool identifying if FastWorkbook should be enabled or disabled</param>
        public void FastWorkbook(bool enable)
        {
            EventManager.FastWorkbook(enable);
        }

        /// <summary>
        /// Sets FastWorksheet custom function, disable some Excel UI functions to improve performance.
        /// </summary>
        /// <param name="ws">An instance of the worksheet which FastWorksheet would be applied to</param>
        /// <param name="enable">Bool identifying if FastWorksheet should be enabled or disabled</param>
        public void FastWorksheet(Excel.Worksheet ws, bool enable)
        {
            EventManager.FastWorksheet(ws, enable);
        }

        /// <summary>
        /// Sets FastWorksheet of all the worksheets in the ActiveWorkbook
        /// </summary>
        /// <param name="enable">Bool identifying if FastWorksheet should be enabled or disabled</param>
        public void FastWorksheets(bool enable)
        {
            FastWorksheets(GetActiveWorkbook(), enable);
        }

        /// <summary>
        /// Sets FastWorksheet of all the worksheets in the selected workbook
        /// </summary>
        /// <param name="wb">An instance of the workbook which the worksheets are contained within</param>
        /// <param name="enable">Bool identifying if FastWorksheet should be enabled or disabled</param>
        public void FastWorksheets(Excel.Workbook wb, bool enable)
        {
            foreach (Excel.Worksheet ws in wb.Sheets)
                FastWorksheet(ws, enable);
        }
    }
}
