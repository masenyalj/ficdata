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
        public static List<IndividualModel> GetIndividual(string filelocation)
        {
            IndividualModel item = new IndividualModel();
            List<IndividualModel> Ind = new List<IndividualModel>();
     
            try
            {
                //Create COM Objects. Create a COM object for everything that is referenced
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filelocation);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets["Sheet1"];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                IndividualModel IndObject = new IndividualModel();


                for (int i = 2; i <= rowCount; i++)
                {
                    // Application ID================

                    if (Convert.ToString(xlRange.Cells[i, 1]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 1].Value2)))
                    {
                        Ind.Add(new IndividualModel() { INDIVIDUAL_Id = xlRange.Cells[i, 1].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { INDIVIDUAL_Id = "-" });
                    }

                    if (Convert.ToString(xlRange.Cells[i, 2]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 2].Value2)))
                    {
                        Ind.Add(new IndividualModel() { TITLE = xlRange.Cells[i, 2].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { TITLE = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 3]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 3].Value2)))
                    {
                        Ind.Add(new IndividualModel() { FIRST_NAME = xlRange.Cells[i, 3].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { FIRST_NAME = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 4]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 4].Value2)))
                    {
                        Ind.Add(new IndividualModel() { SECOND_NAME = xlRange.Cells[i, 4].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { SECOND_NAME = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 5]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 5].Value2)))
                    {
                        Ind.Add(new IndividualModel() { THIRD_NAME = xlRange.Cells[i, 5].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { THIRD_NAME = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 6]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 6].Value2)))
                    {
                        Ind.Add(new IndividualModel() { FOURTH_NAME = xlRange.Cells[i, 6].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { FOURTH_NAME = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 7]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 7].Value2)))
                    {
                        Ind.Add(new IndividualModel() { FullName = xlRange.Cells[i, 7].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { FullName = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 8]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 8].Value2)))
                    {
                        Ind.Add(new IndividualModel() { GENDER = xlRange.Cells[i, 8].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { GENDER = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 9]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 9].Value2)))
                    {
                        Ind.Add(new IndividualModel() { IDNUMBER = xlRange.Cells[i, 9].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { IDNUMBER = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 10]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 10].Value2)))
                    {
                        Ind.Add(new IndividualModel() { PASSPORT = xlRange.Cells[i, 10].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { PASSPORT = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 11]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 11].Value2)))
                    {
                        Ind.Add(new IndividualModel() { DESIGNATION = xlRange.Cells[i, 11].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { DESIGNATION = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 12]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 12].Value2)))
                    {
                        Ind.Add(new IndividualModel() { ListedON = xlRange.Cells[i, 12].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { ListedON = "-" });
                    }

                   

                    if (Convert.ToString(xlRange.Cells[i, 13]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 13].Value2)))
                    {
                        Ind.Add(new IndividualModel() { NATIONALITY = xlRange.Cells[i, 13].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { NATIONALITY = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 14]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 14].Value2)))
                    {
                        Ind.Add(new IndividualModel() { REFERENCE_NUMBER = xlRange.Cells[i, 14].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { REFERENCE_NUMBER = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 15]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 15].Value2)))
                    {
                        Ind.Add(new IndividualModel() { SORT_KEY = xlRange.Cells[i, 15].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { SORT_KEY = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 16]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 16].Value2)))
                    {
                        Ind.Add(new IndividualModel() { SORT_KEY_LAST_MOD = xlRange.Cells[i, 16].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { SORT_KEY_LAST_MOD = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 17]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 17].Value2)))
                    {
                        Ind.Add(new IndividualModel() { SUBMITTED_BY = xlRange.Cells[i, 17].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { SUBMITTED_BY = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 18]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 18].Value2)))
                    {
                        Ind.Add(new IndividualModel() { UN_LIST_TYPE = xlRange.Cells[i, 18].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { UN_LIST_TYPE = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 19]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 19].Value2)))
                    {
                        Ind.Add(new IndividualModel() { COMMENTS = xlRange.Cells[i, 19].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { COMMENTS = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 20]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 20].Value2)))
                    {
                        Ind.Add(new IndividualModel() { VERSIONNUM = xlRange.Cells[i, 20].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { VERSIONNUM = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 21]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 21].Value2)))
                    {
                        Ind.Add(new IndividualModel() { NAME_ORIGINAL_SCRIPT = xlRange.Cells[i, 21].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { NAME_ORIGINAL_SCRIPT = "-" });
                    }





                    /* int name2 = 1;
                       for (int j = 1; j <= colCount; j++)
                       {

                           //new line
                           if (j == 1)
                              Console.Write("\r\n");


                           //write the value to the console
                           if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                                Console.Write(xlRange.Cells[i, j].Value2.ToString() + " ");
                       } */
                }

                return Ind;

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
                Console.WriteLine(ex.Message);
            }
            return Ind;
      
        }
        public static List<EnterpriseModel> GetEnterprise(string filelocation)
        {
            EnterpriseModel item = new EnterpriseModel();
            List<EnterpriseModel> Ind = new List<EnterpriseModel>();

            try
            {
                //Create COM Objects. Create a COM object for everything that is referenced
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(filelocation);
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets["Table1"];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                EnterpriseModel IndObject = new EnterpriseModel();


                for (int i = 2; i <= rowCount; i++)
                {
                    // Application ID================

                    if (Convert.ToString(xlRange.Cells[i, 1]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 1].Value2)))
                    {
                        Ind.Add(new EnterpriseModel() { ENTITY_Id = xlRange.Cells[i, 1].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new EnterpriseModel() { ENTITY_Id = "-" });
                    }

                    if (Convert.ToString(xlRange.Cells[i, 2]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 2].Value2)))
                    {
                        Ind.Add(new EnterpriseModel() { FIRST_NAME = xlRange.Cells[i, 2].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new EnterpriseModel() { FIRST_NAME = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 3]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 3].Value2)))
                    {
                        Ind.Add(new EnterpriseModel() { REFERENCE_NUMBER = xlRange.Cells[i, 3].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new EnterpriseModel() { REFERENCE_NUMBER = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 4]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 4].Value2)))
                    {
                        Ind.Add(new EnterpriseModel() { ListedON = xlRange.Cells[i, 4].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new EnterpriseModel() { ListedON = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 5]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 5].Value2)))
                    {
                        Ind.Add(new EnterpriseModel() { SORT_KEY = xlRange.Cells[i, 5].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new EnterpriseModel() { SORT_KEY = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 6]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 6].Value2)))
                    {
                        Ind.Add(new EnterpriseModel() { SORT_KEY_LAST_MOD = xlRange.Cells[i, 6].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new EnterpriseModel() { SORT_KEY_LAST_MOD = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 7]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 7].Value2)))
                    {
                        Ind.Add(new EnterpriseModel() { SUBMITTEDON = xlRange.Cells[i, 7].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new EnterpriseModel() { SUBMITTEDON = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 8]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 8].Value2)))
                    {
                        Ind.Add(new EnterpriseModel() { UN_LIST_TYPE = xlRange.Cells[i, 8].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new EnterpriseModel() { UN_LIST_TYPE = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 9]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 9].Value2)))
                    {
                        Ind.Add(new EnterpriseModel() { VERSIONNUM = xlRange.Cells[i, 9].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new EnterpriseModel() { VERSIONNUM = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 10]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 10].Value2)))
                    {
                        Ind.Add(new EnterpriseModel() { NAME_ORIGINAL_SCRIPT = xlRange.Cells[i, 10].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new EnterpriseModel() { NAME_ORIGINAL_SCRIPT = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 11]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 11].Value2)))
                    {
                        Ind.Add(new EnterpriseModel() { ISDeleted = xlRange.Cells[i, 11].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new EnterpriseModel() { ISDeleted = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 12]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 12].Value2)))
                    {
                        Ind.Add(new EnterpriseModel() { ApplicationStatus = xlRange.Cells[i, 12].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new EnterpriseModel() { ApplicationStatus = "-" });
                    }



                    if (Convert.ToString(xlRange.Cells[i, 13]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 13].Value2)))
                    {
                        Ind.Add(new EnterpriseModel() { DateInserted = xlRange.Cells[i, 13].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new EnterpriseModel() { DateInserted = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 14]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 14].Value2)))
                    {
                        Ind.Add(new EnterpriseModel() { COMMENTS = xlRange.Cells[i, 14].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new EnterpriseModel() { COMMENTS = "-" });
                    }


                    if (Convert.ToString(xlRange.Cells[i, 15]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 15].Value2)))
                    {
                        Ind.Add(new EnterpriseModel() { NOTE = xlRange.Cells[i, 15].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new EnterpriseModel() { NOTE = "-" });
                    }

                    

                }

                return Ind;

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
                Console.WriteLine(ex.Message);
            }
            return Ind;
           
        }
        public static List<IndividualModel> ReadData(string filelocation, int sheetname)
        {
            List<string> numberList = new List<string>();
           // IEnumerable<IndividualModel> IndvObject = new IEnumerable<IndividualModel>();
          //  List<IndividualModel> list = new List<IndividualModel>();
            IndividualModel item = new IndividualModel();
            List<IndividualModel> Ind = new List<IndividualModel>();
           // IndividualModelset();

            // item.ApplicationStatus = 

            try
            {
                //Create COM Objects. Create a COM object for everything that is referenced
                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\lmasenya\Documents\News.xlsx");
                Excel._Worksheet xlWorksheet = xlWorkbook.Sheets["Table"];
                Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                IndividualModel IndObject = new IndividualModel();
            

                for (int i = 2; i <= rowCount; i++)
                {

                    if (Convert.ToString(xlRange.Cells[i, 1]) != null && !String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, 1].Value2)))
                    {
                        Ind.Add(new IndividualModel() { INDIVIDUAL_Id = xlRange.Cells[i, 1].Value2.ToString() });
                    }
                    else
                    {
                        Ind.Add(new IndividualModel() { INDIVIDUAL_Id = "-" });
                    }
                    



                    /* int name2 = 1;
                       for (int j = 1; j <= colCount; j++)
                       {

                           //new line
                           if (j == 1)
                              Console.Write("\r\n");


                           //write the value to the console
                           if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                                Console.Write(xlRange.Cells[i, j].Value2.ToString() + " ");
                       } */
                }

                return Ind;

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
                Console.WriteLine(ex.Message);
            }
            return Ind;
        }
    }
}
