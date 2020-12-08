using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using CIPC.FICData;
using IronXL;
using System.Linq;

namespace CIPC.FICData
{
    public class FICData
    {
        public static void downloadficxcel(string filelocation, string saveto)
        {
            try
            {
                WebClient webClient = new WebClient();
                webClient.DownloadFile(filelocation, saveto);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        public static List<string> ReadData(string filelocation, int sheetname)
        {
            List<string> numberList = new List<string>();

            try
            {
                //Create COM Objects. Create a COM object for everything that is referenced
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filelocation);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[sheetname];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                
                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {

                        //new line
                        if (j == 1)
                           numberList.Add("\r\n");

                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                              numberList.Add(xlRange.Cells[i, j].Value2.ToString() + " ");
                    }
                }
                return numberList;
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:
                //  never use two dots, all COM objects must be referenced and released individually
                //  ex: [somthing].[something].[something] is bad

                //release com objects to fully kill excel process from running in the background
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
            catch (Exception ex)
            {
                //return ex.Message;
            }
            return numberList;
        }
    }
}
