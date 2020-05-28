using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;

namespace CreateExcelSpreadSheet
{
    public delegate Worksheet TestDelegate(Worksheet ExcelWorkSheet);

    internal class Program
    {
        public static TestDelegate Sheet1 = DelClass.AddContent1;
        public static TestDelegate Sheet2 = DelClass.AddContent2;

        private static void Main(string[] args)
        {
            Application ExcelApp = new Application();
            Workbook ExcelWorkBook = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            //Workbook ExcelWorkBook = null;
            ExcelApp.Visible = true;

            try
            {
                ExcelWorkBook = AddWorkSheet(ExcelWorkBook, Sheet1);
                ExcelWorkBook = AddWorkSheet(ExcelWorkBook, Sheet2);
                ExcelWorkBook.SaveAs($@"C:\Users\Admin\Desktop\Testing{DateTime.Now:yyyyMMddhhmmss}.xlsx");
                ExcelWorkBook.Close();
                ExcelApp.Quit();

                Marshal.ReleaseComObject(ExcelWorkBook);
                Marshal.ReleaseComObject(ExcelApp);
            }
            catch (Exception exHandle)
            {
                Console.WriteLine("Exception: " + exHandle.Message);
                Console.ReadLine();
            }
            finally
            {
                foreach (Process process in Process.GetProcessesByName("Excel"))
                {
                    process.Kill();
                }
            }
        }

        public static Workbook AddWorkSheet(Workbook ExcelWorkBook, TestDelegate testDelegate)
        {
            Worksheet ExcelWorkSheet = (Worksheet)ExcelWorkBook.Sheets.Add();
            ExcelWorkSheet = testDelegate(ExcelWorkSheet);

            Marshal.ReleaseComObject(ExcelWorkSheet);
            return ExcelWorkBook;
        }
    }

    public class DelClass
    {
        public static Worksheet AddContent1(Worksheet ExcelWorkSheet)
        {
            ExcelWorkSheet.Name = "Sheet2";
            for (int r = 1; r < 3; r++)
            {
                for (int c = 1; c < 3; c++)
                {
                    ExcelWorkSheet.Cells[r, c] = $"R{r}C{c}";
                }
            }
            return ExcelWorkSheet;
        }

        public static Worksheet AddContent2(Worksheet ExcelWorkSheet)
        {
            ExcelWorkSheet.Name = "Sheet3";
            for (int r = 1; r < 3; r++)
            {
                for (int c = 1; c < 3; c++)
                {
                    ExcelWorkSheet.Cells[r, c] = $"R{r}C{c}";
                }
            }
            return ExcelWorkSheet;
        }
    }
}