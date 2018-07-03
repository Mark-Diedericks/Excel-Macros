/*
 * Mark Diedericks
 * 02/07/2015
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
        //Get the excel application
        public Excel.Application GetApplication()
        {
            return Main.GetApplication();
        }

        //Get the active workbook
        public Excel.Workbook GetActiveWorkbook()
        {
            return GetApplication().ActiveWorkbook;
        }

        //Get the active worksheet
        public Excel.Worksheet GetActiveWorksheet()
        {
            return (Excel.Worksheet)GetActiveWorkbook().ActiveSheet;
        }

        //Get a range selection through the excel inputbox
        public Excel.Range RequestRangeInput(string message)
        {
            object result = MessageManager.DisplayInputMessage(message, "Input Range", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8);
            Main.FireFocusEvent();
            return result as Excel.Range;
        }

        //Get a boolean input through a message box
        public bool RequestBooleanInput(string message)
        {
            bool result = MessageManager.DisplayYesNoMessage(message, "Boolean Input");
            Main.FireFocusEvent();
            return result;
        }

        //Display a message in a message box
        public void DisplayMessage(string message, string caption)
        {
            MessageManager.DisplayOkMessage(message, caption);
            Main.FireFocusEvent();
        }

        public void FastWorkbook(bool enable)
        {
            EventManager.FastWorkbook(enable);
        }

        public void FastWorksheet(Excel.Worksheet ws, bool enable)
        {
            EventManager.FastWorksheet(ws, enable);
        }

        //Make all the worksheets in the active workbook fast
        public void FastWorksheets(bool enable)
        {
            FastWorksheets(GetActiveWorkbook(), enable);
        }

        //Make all the worksheets in a selected workbook fast
        public void FastWorksheets(Excel.Workbook wb, bool enable)
        {
            foreach (Excel.Worksheet ws in wb.Sheets)
                FastWorksheet(ws, enable);
        }
    }
}
