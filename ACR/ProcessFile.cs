using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ACR
{
    class ProcessFile
    {
        private static string file = @"Import\100day.csv";

        private string[] field1;
        private string[] field2;
        private string[] field3;
        private string[] field4;
        private string[] field5;
        private string[] field6;
        private string[] field7;
        private string[] field8;
        private string[] field9;
        private string[] field10;
        private string[] field11;
        private string[] field12;
        private string[] field13;
        private string[] field14;
        private string[] field15;
        private string[] field16;
        private string[] field17;
        private string[] field18;
        private string[] field19;
        private string[] field20;

        public ProcessFile(bool x)
        {
            if (x == false)
            {
                Environment.Exit(0);
            }
        }

        public ProcessFile(int length)
        {
            field1 = new string[length];
            field2 = new string[length];
            field3 = new string[length];
            field4 = new string[length];
            field5 = new string[length];
            field6 = new string[length];
            field7 = new string[length];
            field8 = new string[length];
            field9 = new string[length];
            field10 = new string[length];
            field11 = new string[length];
            field12 = new string[length];
            field13 = new string[length];
            field14 = new string[length];
            field15 = new string[length];
            field16 = new string[length];
            field17 = new string[length];
            field18 = new string[length];
            field19 = new string[length];
            field20 = new string[length];
        }

        public void RawRead()
        {
            // Example #2
            // Read each line of the file into a string array. Each element
            // of the array is one line of the file.
            string[] rawData = System.IO.File.ReadAllLines(file);

            // Display the file contents by using a foreach loop.            
            int count = 0;
            foreach (string rawRecord in rawData)
            {
                //Console.WriteLine("\t" + rawRecord);     
                string[] temp = rawRecord.Split('|');
                field1[count] = temp[0];
                field2[count] = temp[1];
                field3[count] = temp[2];
                field4[count] = temp[3];
                field5[count] = temp[4];
                field6[count] = temp[5];
                field7[count] = temp[6];
                field8[count] = temp[7];
                field9[count] = temp[8];
                field10[count] = temp[9];
                field11[count] = temp[10];
                field12[count] = temp[11];
                field13[count] = temp[12];
                field14[count] = temp[13];
                field15[count] = temp[14];
                field16[count] = temp[15];
                field17[count] = temp[16];
                field18[count] = temp[17];
                field19[count] = temp[18];
                field20[count] = temp[19];
                count++;
            }

            Console.WriteLine("\nRaw file import done \nDo you want a preview output? (Y/N)");
            ConsoleKeyInfo read = Console.ReadKey();
            if (read.KeyChar.ToString() == "y" || read.KeyChar.ToString() == "Y")
            {
                Console.WriteLine("\nHow many rows do you want to preview? ");
                ConsoleKeyInfo previewAmt = Console.ReadKey();
                try
                {
                    previewOutput(Int32.Parse(previewAmt.KeyChar.ToString()));
                }
                catch
                {
                    Console.Clear();
                    Console.WriteLine("Error occured...\nTry import again? (Y/N)");
                    ConsoleKeyInfo errorRead = Console.ReadKey();
                    if (errorRead.KeyChar.ToString() == "y" || errorRead.KeyChar.ToString() == "Y")
                    {
                        reRunRawImport();
                    }
                }
            }
            startExcelProcess();
        }

        private void previewOutput(int previewAmt)
        {
            if (previewAmt > (field1.Length))
            {
                previewAmt = field1.Length;
            }

            for (int y = 0; y < previewAmt; y++)
            {
                if (y == 0)
                {
                    Console.WriteLine(String.Format("\n\nPREVIEW OUTPUT:\n\n{0,-20} | {1,-20} | {2,-20} | {3,-20} | {4,-20} | {5,-20} | {6,-20} | {7,-20} | {8,-20} | {9,-20}", "Field 1", "Field 2", "Field 3", "Field 4", "Field 5", "Field 6", "Field 7", "Field 8", "Field 9", "Field 10"));
                    Console.WriteLine("--------------------------------------------------------------------------------");
                }

                Console.WriteLine(String.Format("{0,-20} | {1,-20} | {2,-20} | {3,-20} | {4,-20} | {5,-20} | {6,-20} | {7,-20} | {8,-20} | {9,-20}", field1[y], field2[y], field3[y], field4[y], field5[y], field6[y], field7[y], field8[y], field8[y], field10[y]));
            }
            Console.WriteLine("--------------------------------------------------------------------------------\n");

            Console.WriteLine("Do you need to run the raw import again? (Y/N)");
            ConsoleKeyInfo read = Console.ReadKey();
            if (read.KeyChar.ToString() == "y" || read.KeyChar.ToString() == "Y")
            {
                reRunRawImport();
            }
        }

        private void reRunRawImport()
        {
            Console.Clear();
            Program pro = new Program();
            pro.reRun();
        }

        public string getField1Value(int index)
        {
            string value = field1[index];
            return value;
        }

        public string getField2Value(int index)
        {
            string value = field2[index];
            return value;
        }

        public string getField3Value(int index)
        {
            string value = field3[index];
            return value;
        }

        public string getField4Value(int index)
        {
            string value = field4[index];
            return value;
        }

        //AMS - Who reports to who (AMS- Nickie vs AMS-Jimmy Excel Worksheets)        
        private static string AMS_ReportsTo_File = @"Import\AMS_ReportsTo.csv";
        static int AMS_Length = File.ReadLines(AMS_ReportsTo_File).Count();
        string[,] ReportsTo = new string[AMS_Length, 2];

        private void AMS_ReportsTo_Create()
        {
            string[] rawData1 = System.IO.File.ReadAllLines(AMS_ReportsTo_File);
            int count = 0;
            foreach (string rawRecord in rawData1)
            {
                string[] temp = rawRecord.Split(',');
                ReportsTo[count, 0] = temp[0];
                ReportsTo[count, 1] = temp[1];
                count++;
            }
        }

        private string AMS_ReportsTo_Lookup(string scc)
        {
            //search array for SCC. Lookup 1:2
            int count = 0;
            while (count < (ReportsTo.Length / 2))
            {
                if (scc.Contains(ReportsTo[count, 0]))
                {
                    return ReportsTo[count, 1];
                }
                count++;
            }
            return "fail";
        }

        private string SCC_Worksheet_Translation(string scc, string group)
        {
            //switch (group)
            //{
            //    case string a when a.Contains("GWFM Client Services-T1-AMS-ADPGHCM"):
            //        if (AMS_ReportsTo_Lookup(scc) == "N")
            //            return "AMS-Nickie";
            //        else if (AMS_ReportsTo_Lookup(scc) == "J")
            //            return "AMS-Jimmy";
            //        break;
            //    case string b when b.Contains("GWFM Client Services-T1-APAC-ADPGHCM"):
            //        return "APAC";                    
            //    case string c when c.Contains("GWFM Client Services-T1-EMEA-BCN-ADPGHCM"):
            //        return "EMEA-BCN";
            //    case string d when d.Contains("GWFM Client Services-T1-EMEA-PRG-ADPGHCM"):
            //        return "EMEA-PRG";
            //    default: //do nothing
            //        break;
            //}

            if (group.Contains("GWFM Client Services-T1-AMS-ADPGHCM"))
            {
                if (AMS_ReportsTo_Lookup(scc) == "N")
                    return "AMS-Nickie";
                else if (AMS_ReportsTo_Lookup(scc) == "J")
                    return "AMS-Jimmy";
            }
            else if (group.Contains("GWFM Client Services-T1-APAC-ADPGHCM"))
            {
                return "APAC";
            }
            else if (group.Contains("GWFM Client Services-T1-EMEA-BCN-ADPGHCM"))
            {
                return "EMEA-BCN";
            }
            else if (group.Contains("GWFM Client Services-T1-EMEA-PRG-ADPGHCM"))
            {
                return "EMEA-PRG";
            }

            return "fail";
        }

        private void startExcelProcess()
        {
            // Get the current date.
            DateTime PCDATE = DateTime.Today;

            string outputFile = @"Output\100_Day_GWFM_" + PCDATE.ToString("yyyyMMdd") + ".xlsx";
            File.Delete(outputFile);
            FileInfo excelInfo = new FileInfo(outputFile);
            ExcelPackage excel = new ExcelPackage(excelInfo);

            //temp array to make worksheets
            ArrayList worksheetArray = new ArrayList();

            //Creates the Worksheets (One per region)
            AMS_ReportsTo_Create(); //Needed to figure out which AMS page SCC will fall on
            worksheetArray.Add("AMS-Nickie");
            worksheetArray.Add("AMS-Jimmy");
            worksheetArray.Add("APAC");
            worksheetArray.Add("EMEA-BCN");
            worksheetArray.Add("EMEA-PRG");



            //worksheetArray.RemoveAt(0);
            //worksheetArray.Sort();

            for (int i = 0; i < worksheetArray.Count; i++)
            {
                var worksheetNM = excel.Workbook.Worksheets.Add(worksheetArray[i].ToString());
                worksheetNM.Cells["A1"].Value = field8[0];  //Owner
                worksheetNM.Cells["B1"].Value = field1[0];  //SR#
                worksheetNM.Cells["C1"].Value = field19[0]; //SR Age
                worksheetNM.Cells["D1"].Value = field12[0]; //Date Created
                worksheetNM.Cells["E1"].Value = field3[0];  //Priority
                worksheetNM.Cells["F1"].Value = field4[0];  //Summary
                worksheetNM.Cells["G1"].Value = field6[0];  //Account
                worksheetNM.Cells["H1"].Value = field11[0]; //Last Name
                worksheetNM.Cells["I1"].Value = field10[0]; //First Name
                worksheetNM.Cells["J1"].Value = "NOTES                                                                                                                |";


                int row = 2;
                for (int x = 0; x < field14.Length; x++)
                {
                    if (SCC_Worksheet_Translation(field8[x], field14[x]).Contains(worksheetNM.ToString()))
                    {
                        //MAKES BODY
                        worksheetNM.Cells["A" + row.ToString()].Value = field8[x];
                        worksheetNM.Cells["B" + row.ToString()].Value = field1[x];
                        worksheetNM.Cells["C" + row.ToString()].Value = field19[x];
                        worksheetNM.Cells["D" + row.ToString()].Value = field12[x];
                        worksheetNM.Cells["E" + row.ToString()].Value = field3[x];
                        worksheetNM.Cells["F" + row.ToString()].Value = field4[x];
                        worksheetNM.Cells["G" + row.ToString()].Value = field6[x];
                        worksheetNM.Cells["H" + row.ToString()].Value = field11[x];
                        worksheetNM.Cells["I" + row.ToString()].Value = field10[x];

                        //HEADER BOARDER
                        worksheetNM.Cells["A1"].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["A1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["A1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["A1"].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["B1"].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["B1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["B1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["B1"].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["C1"].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["C1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["C1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["C1"].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["D1"].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["D1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["D1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["D1"].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["E1"].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["E1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["E1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["E1"].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["F1"].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["F1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["F1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["F1"].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["G1"].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["G1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["G1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["G1"].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["H1"].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["H1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["H1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["H1"].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["I1"].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["I1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["I1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["I1"].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["J1"].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["J1"].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["J1"].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["J1"].Style.Border.Left.Style = ExcelBorderStyle.Medium;

                        //THE BOADER FOR THE BODY
                        worksheetNM.Cells["A" + row.ToString()].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["A" + row.ToString()].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["A" + row.ToString()].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["A" + row.ToString()].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["B" + row.ToString()].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["B" + row.ToString()].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["B" + row.ToString()].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["B" + row.ToString()].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["C" + row.ToString()].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["C" + row.ToString()].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["C" + row.ToString()].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["C" + row.ToString()].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["D" + row.ToString()].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["D" + row.ToString()].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["D" + row.ToString()].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["D" + row.ToString()].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["E" + row.ToString()].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["E" + row.ToString()].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["E" + row.ToString()].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["E" + row.ToString()].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["F" + row.ToString()].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["F" + row.ToString()].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["F" + row.ToString()].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["F" + row.ToString()].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["G" + row.ToString()].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["G" + row.ToString()].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["G" + row.ToString()].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["G" + row.ToString()].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["H" + row.ToString()].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["H" + row.ToString()].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["H" + row.ToString()].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["H" + row.ToString()].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["I" + row.ToString()].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["I" + row.ToString()].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["I" + row.ToString()].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["I" + row.ToString()].Style.Border.Left.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["J" + row.ToString()].Style.Border.Top.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["J" + row.ToString()].Style.Border.Right.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["J" + row.ToString()].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
                        worksheetNM.Cells["J" + row.ToString()].Style.Border.Left.Style = ExcelBorderStyle.Medium;

                        row++;
                    }
                }

                worksheetNM.Cells[1, 1, 1, 10].Style.Font.Bold = true;
                Color colFromHex = ColorTranslator.FromHtml("255,255,0");
                worksheetNM.Cells[1, 1, 1, 10].Style.Fill.PatternType = ExcelFillStyle.Solid;
                worksheetNM.Cells[1, 1, 1, 10].Style.Fill.BackgroundColor.SetColor(colFromHex);

                worksheetNM.Cells[worksheetNM.Dimension.Address].AutoFitColumns();
                worksheetNM.Column(6).Width = 75;
                worksheetNM.Column(10).Style.WrapText = true;

                worksheetNM.Column(1).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheetNM.Column(2).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheetNM.Column(3).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheetNM.Column(4).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheetNM.Column(5).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheetNM.Column(6).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheetNM.Column(7).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheetNM.Column(8).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheetNM.Column(9).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                worksheetNM.Column(10).Style.VerticalAlignment = ExcelVerticalAlignment.Top;
            }

            excel.Save();

            Console.WriteLine("\n\nReport can be found in the following folder: " + outputFile);
            File.Delete(file);
            Console.ReadKey();
        }
    }
}
