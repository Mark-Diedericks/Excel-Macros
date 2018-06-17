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
        private Excel.Application GetApplication()
        {
            return Main.GetApplication();
        }

        //Get the active workbook
        private Excel.Workbook GetActiveWorkbook()
        {
            return GetApplication().ActiveWorkbook;
        }

        //Get the active worksheet
        private Excel.Worksheet GetActiveWorksheet()
        {
            return (Excel.Worksheet)GetActiveWorkbook().ActiveSheet;
        }

        //Get a range selection through the excel inputbox
        public Excel.Range RequestRangeInput(string message)
        {
            Main.GetExcelDispatcher().Invoke(() => Main.SetExcelInteractive(true));

            Excel.Range res = GetApplication().InputBox(message, "Input Range", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, 8) as Excel.Range;

            Main.GetExcelDispatcher().Invoke(() => Main.SetExcelInteractive(false));
            
            Main.FireFocusEvent();

            return res;
        }

        //Get a boolean input through a message box
        public bool RequestBooleanInput(string message)
        {
            Main.GetExcelDispatcher().Invoke(() => Main.SetExcelInteractive(true));

            bool res = MessageBox.Show(message, "Boolean Input", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;

            Main.GetExcelDispatcher().Invoke(() => Main.SetExcelInteractive(false));
            
            Main.FireFocusEvent();

            return res;
        }

        //Display a message in a message box
        public void DisplayMessage(string message, string caption)
        {
            Main.GetExcelDispatcher().Invoke(() => Main.SetExcelInteractive(true));

            MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);

            Main.GetExcelDispatcher().Invoke(() => Main.SetExcelInteractive(false));

            Main.FireFocusEvent();
        }

        //
        //Generic Functions
        //

        //FastWorksheet Macro by Paul Bica
        //https://stackoverflow.com/questions/30959315/excel-vba-performance-1-million-rows-delete-rows-containing-a-value-in-less
        public void FastWorkbook(bool enable)
        {
            GetApplication().Calculation = enable ? Excel.XlCalculation.xlCalculationManual : Excel.XlCalculation.xlCalculationAutomatic;
            GetApplication().DisplayAlerts = !enable;
            GetApplication().DisplayStatusBar = !enable;
            GetApplication().EnableAnimations = !enable;
            GetApplication().EnableEvents = !enable;
            GetApplication().ScreenUpdating = !enable;

            FastWorksheets(enable); //Make all worksheets in the active workbook fast
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

        //FastWorksheet Macro by Paul Bica
        //https://stackoverflow.com/questions/30959315/excel-vba-performance-1-million-rows-delete-rows-containing-a-value-in-less
        public void FastWorksheet(Excel.Worksheet ws, bool enable)
        {
            ws.DisplayPageBreaks = false;
            ws.EnableCalculation = !enable;
            ws.EnableFormatConditionsCalculation = !enable;
            ws.EnablePivotTable = !enable;
        }
    }
}
