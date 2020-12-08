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
        public static object ReadData(string filelocation, string sheetname, string cellName, string strRange)
        {
            try
            {
                WorkBook workbook = WorkBook.Load(filelocation);
                WorkSheet sheet = workbook.GetWorkSheet(sheetname);
                //Select cells easily in Excel notation and return the calculated value
                int cellValue = sheet[cellName].IntValue;
                // Read from Ranges of cells elegantly.
                foreach (var cell in sheet[strRange])
                {
                    return cell.Value;
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            return "No data";
        }
    }
}
