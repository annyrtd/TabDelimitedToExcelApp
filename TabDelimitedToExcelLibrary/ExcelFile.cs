using System;
using Microsoft.Office.Interop.Excel;

namespace TabDelimitedToExcelLibrary
{
    public class ExcelFile
    {
        Application xlApp;
        object misValue;
        Workbook xlWorkBook;
        Worksheet xlWorkSheet;

        public Range Cells { get; set; }

        public ExcelFile()
        {
            xlApp = new Application();
            misValue = System.Reflection.Missing.Value;
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);
            Cells = xlWorkSheet.Cells;
        }

        public void SaveAs(string newFileName)
        {
            xlApp.DisplayAlerts = false;
            xlWorkBook.SaveAs(newFileName,
                     XlFileFormat.xlWorkbookNormal,
                     misValue,
                     misValue,
                     misValue,
                     misValue,
                     XlSaveAsAccessMode.xlNoChange,
                     misValue,
                     misValue,
                     misValue,
                     misValue,
                     misValue);           
        }

        public void Close()
        {
            xlWorkBook.Close(false, misValue, misValue);
            xlApp.Quit();
            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception)
            {
                obj = null;
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}