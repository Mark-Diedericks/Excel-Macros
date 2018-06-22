/*
 * Mark Diedericks
 * 09/06/2015
 * Version 1.0.0
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
        public void RequestRangeInput(string message, Action<Excel.Range> OnResult)
        {
            MessageManager.DisplayInputMessage(message, "Input Range", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8, (res) =>
            {
                Excel.Range range = res as Excel.Range;
                Main.FireFocusEvent();
                OnResult?.Invoke(range);
            });
        }

        //Get a boolean input through a message box
        public void RequestBooleanInput(string message, Action<bool> OnResult)
        {
            MessageManager.DisplayYesNoMessage(message, "Boolean Input", (res) =>
            {
                Main.FireFocusEvent();
                OnResult?.Invoke(res);
            });
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
