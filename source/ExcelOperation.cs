using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Drawing;
using ImageProcessor.Processors;

namespace Reconciliation
{
    internal static class ExcelOperation
    {
        public static void ExcelReconciliation(List<AccountRow> accounts, string strExcelPath)
        {
            // Create a new Excel file
            Application excelApp = new Application();
            Workbook workbook = excelApp.Workbooks.Open(strExcelPath);

            try
            {
                Worksheet accBookSheet = (Worksheet)workbook.ActiveSheet;
                Worksheet bankSheet = null;
                Worksheet consolidated = null;
                Worksheet summary = null;
                // Add a new sheet to the workbook
                foreach (Worksheet item in workbook.Worksheets)
                {
                    if (string.Equals(item.Name, "Bank Statement", StringComparison.OrdinalIgnoreCase))
                    {
                        bankSheet = item;
                    }
                    else if (string.Equals(item.Name, "Consolidated", StringComparison.OrdinalIgnoreCase))
                    {
                        consolidated = item;
                    }
                    else if (string.Equals(item.Name, "Summary", StringComparison.OrdinalIgnoreCase))
                    {
                        summary = item;
                    }
                    else
                    {
                        accBookSheet = item;
                        accBookSheet.Activate();
                    }
                }
                if (bankSheet != null)
                {
                    bankSheet.Cells.ClearContents();
                }
                else
                {
                    bankSheet = workbook.Sheets.Add();
                    bankSheet.Name = "Bank Statement";
                }

                if (consolidated != null)
                {
                    consolidated.Cells.ClearContents();
                }
                else
                {
                    consolidated = workbook.Sheets.Add();
                    consolidated.Name = "Consolidated";
                }

                if (summary != null)
                {
                    summary.Cells.ClearContents();
                }
                else
                {
                    summary = workbook.Sheets.Add();
                    summary.Name = "Summary";
                }
                Range column1Range = bankSheet.Range["A:A"];

                // Set the format of the cells in column 1 to "Text"
                column1Range.NumberFormat = "@";
                bankSheet.Cells[1, 1] = "Date";
                bankSheet.Cells[1, 2] = "Money In (Cr.)";
                bankSheet.Cells[1, 3] = "Money Out (Dr.)";
                bankSheet.Cells[1, 4] = "Balance";
                // Write the tabular data to the worksheet
                for (int i = 0; i < accounts.Count; i++)
                {
                    bankSheet.Cells[i + 2, 1] = accounts[i].Date;
                    if (accounts[i].Deposit != 0)
                        bankSheet.Cells[i + 2, 2] = accounts[i].Deposit;
                    if (accounts[i].Withdrawal != 0)
                        bankSheet.Cells[i + 2, 3] = accounts[i].Withdrawal;
                    bankSheet.Cells[i + 2, 4] = accounts[i].Balance;

                }
                // Auto-fit all columns
                Range range = bankSheet.UsedRange;
                range.Columns.AutoFit();


                // Get the range of the cells you want to merge
                Range rangeToMerge1 = consolidated.Range["B1:C1"];
                rangeToMerge1.Merge();
                rangeToMerge1.Value = "Bank Statement";



                Range destinationStartCell = consolidated.Cells[2, 1];
                // Copy the source range to the destination range
                Range consolidatedRange = consolidated.Range[destinationStartCell.Address];
                range.Copy(consolidatedRange);
                // get the range of the column to remove
                Range rangeDelete = (Range)consolidated.Columns[4];
                // delete the entire column
                rangeDelete.Delete();

                Range rangeToMerge2 = consolidated.Range["D1:E1"];
                rangeToMerge2.Merge();
                rangeToMerge2.Value = "Accounting Books";
                consolidated.Cells[2, 4] = "Money In (Cr.)";
                consolidated.Cells[2, 5] = "Money Out (Dr.)";
                consolidated.Cells[2, 6] = "Description";

                Range accountRange = accBookSheet.UsedRange;

                Dictionary<string, List<Book>> bookData = new Dictionary<string, List<Book>>();
                accBookSheet.Activate();
                for (int row = 2; row <= accountRange.Rows.Count; row++)
                {

                    string dt = accBookSheet.getValueDT(row, 1);
                    if (dt == default(string))
                        break;
                    if (bookData.ContainsKey(dt))
                    {
                        bookData[dt].Add(new Book(accBookSheet.getValue(row, 4), accBookSheet.getValue(row, 3), accBookSheet.getValue(row, 2)));
                    }
                    else
                    {
                        bookData.Add(dt, new List<Book>() { new Book(accBookSheet.getValue(row, 4), accBookSheet.getValue(row, 3), accBookSheet.getValue(row, 2)) });
                    }
                }

                // Iterate over each row in the used range and search for the date
                for (int crow = 3; crow <= consolidated.UsedRange.Rows.Count; crow++)
                {
                    string bookWithdrawal = consolidated.getValue(crow, 3);
                    string bookDeposit = consolidated.getValue(crow, 2);
                    string bookValue = bookWithdrawal;
                    if (bookWithdrawal == null || bookWithdrawal == "")
                        bookValue = bookDeposit;

                    string cDate = consolidated.getValueDT(crow, 1);
                    if (cDate == default(string))
                        break;
                    for (int i = 0; i < bookData.Count; i++)
                    {
                        string bDt = bookData.Keys.ToList()[i];
                        if (bDt == cDate)
                        {
                            bool found = false;
                            for (int z = 0; z < bookData[bDt].Count; z++)
                            {
                                Book b = bookData[bDt][z];
                                if (b.BookWithdrawal != "" && b.BookWithdrawal != null && bookWithdrawal == b.BookWithdrawal)
                                {
                                    consolidated.Cells[crow, 5] = bookWithdrawal;
                                    consolidated.Cells[crow, 6] = b.BookDescription;
                                    found = true;
                                    bookData[bDt].Remove(b);
                                    z--;

                                    //if (bookDeposit == bookValue)
                                    //{
                                    //    //((Range)consolidated.Cells[crow, 5]).Font.Color = ColorTranslator.ToOle(Color.Turquoise);
                                    //    ((Range)consolidated.Cells[crow, 7]).Interior.Color = ColorTranslator.ToOle(Color.Turquoise);
                                    //}
                                    break;
                                }
                                else if (b.BookDeposit != "" && b.BookDeposit != null && bookDeposit == b.BookDeposit)
                                {
                                    consolidated.Cells[crow, 4] = bookValue;
                                    consolidated.Cells[crow, 6] = b.BookDescription;
                                    found = true;
                                    bookData[bDt].Remove(b);
                                    //if (bookWithdrawal == bookValue)
                                    //{
                                    //    //((Range)consolidated.Cells[crow, 4]).Font.Color = ColorTranslator.ToOle(Color.Turquoise);
                                    //    ((Range)consolidated.Cells[crow, 7]).Interior.Color = ColorTranslator.ToOle(Color.Turquoise);
                                    //}
                                    z--;
                                    break;
                                }
                            }
                            if (found)
                            {
                                break;
                            }
                        }
                    }
                }


                Range copyConso = consolidated.Range["A1", "F2"];

                Range destSummary = summary.Range["A2"];

                // Copy the range to the destination range
                copyConso.Copy(destSummary);
                summary.Range["A:A"].NumberFormat = "@";

                bool passed = true;
                int rowSummary = 4;
                // Iterate over each row in the used range and search for the date
                for (int crow = 3; crow <= consolidated.UsedRange.Rows.Count; crow++)
                {
                    string bankWithdrawal = consolidated.getValue(crow, 3);
                    string bankDeposit = consolidated.getValue(crow, 2);
                    string bookWithdrawal = consolidated.getValue(crow, 5);
                    string bookDeposit = consolidated.getValue(crow, 4);
                    string bankValue = bankWithdrawal;
                    if (bankWithdrawal == null || bankWithdrawal == "")
                        bankValue = bankDeposit;
                    string bookValue = bookWithdrawal;
                    if (bookWithdrawal == null || bookWithdrawal == "")
                        bookValue = bookDeposit;

                    if (bookValue == null || bookValue == "")
                    {
                        passed = false;
                        ((Range)consolidated.Cells[crow, 1]).Font.Color = ColorTranslator.ToOle(Color.Red);
                        ((Range)consolidated.Cells[crow, 2]).Font.Color = ColorTranslator.ToOle(Color.Red);
                        ((Range)consolidated.Cells[crow, 3]).Font.Color = ColorTranslator.ToOle(Color.Red);
                        ((Range)consolidated.Cells[crow, 6]).Font.Color = ColorTranslator.ToOle(Color.Red);
                        ((Range)consolidated.Cells[crow, 7]).Interior.Color = ColorTranslator.ToOle(Color.Red);
                        string cDate = consolidated.getValueDT(crow, 1);
                        if (cDate == default(string))
                            break;

                        for (int i = 0; i < bookData.Count; i++)
                        {
                            string bDt = bookData.Keys.ToList()[i];
                            if (bDt == cDate)
                            {

                                List<double> lstdouble = new List<double>();
                                Dictionary<int, double> dictIndex = new Dictionary<int, double>();
                                for (int z = 0; z < bookData[bDt].Count; z++)
                                {
                                    Book b = bookData[bDt][z];
                                    if (b.BookWithdrawal != null && b.BookWithdrawal != "")
                                    {
                                        lstdouble.Add(Convert.ToDouble(b.BookWithdrawal));
                                        dictIndex.Add(z, lstdouble[z]);
                                    }
                                    else if (b.BookDeposit != null && b.BookDeposit != "")
                                    {
                                        lstdouble.Add(Convert.ToDouble(b.BookDeposit));
                                        dictIndex.Add(z, lstdouble[z]);
                                    }


                                }
                                if (lstdouble.Count < 1)
                                    continue;
                                double closestDouble = lstdouble.OrderBy(d => Math.Abs(d - Convert.ToDouble(bankValue))).First();
                                int index = dictIndex.GetKeyFromValue(closestDouble);
                                if (index != -1)
                                {
                                    consolidated.Cells[crow, 6] = bookData[bDt][index].BookDescription;
                                   

                                    string closetString = closestDouble.ToString();;

                                    if (closetString == bookData[bDt][index].BookWithdrawal)
                                    {
                                        consolidated.Cells[crow, 5] = closetString;
                                        ((Range)consolidated.Cells[crow, 5]).Font.Color = ColorTranslator.ToOle(Color.Red);
                                    }
                                    else
                                    {
                                        consolidated.Cells[crow, 4] = closetString;
                                        ((Range)consolidated.Cells[crow, 4]).Font.Color = ColorTranslator.ToOle(Color.Red);
                                    }
                                    bookData[bDt].RemoveAt(index);
                                }

                            }
                        }

                        summary.Cells[rowSummary, 1] = consolidated.Cells[crow, 1];
                        summary.Cells[rowSummary, 2] = consolidated.Cells[crow, 2];
                        summary.Cells[rowSummary, 3] = consolidated.Cells[crow, 3];
                        summary.Cells[rowSummary, 4] = consolidated.Cells[crow, 4];
                        summary.Cells[rowSummary, 5] = consolidated.Cells[crow, 5];
                        summary.Cells[rowSummary, 6] = consolidated.Cells[crow, 6];

                        rowSummary++;
                    }
                }

                if(passed)
                {
                    summary.Cells[1, 1] = "Reconciliation";
                    summary.Cells[1, 2] = "Passed";
                    ((Range)summary.Cells[1, 2]).Font.Color = ColorTranslator.ToOle(Color.Green);
                }
                else
                {
                    summary.Cells[1, 1] = "Reconciliation";
                    summary.Cells[1, 2] = "Failed";
                    ((Range)summary.Cells[1, 2]).Font.Color = ColorTranslator.ToOle(Color.Red);
                }

                consolidated.UsedRange.Columns.AutoFit();
                summary.UsedRange.Columns.AutoFit();
                summary.Activate();
                // Save the Excel file
                //string fileName = Path.Combine(Path.GetDirectoryName(strPDFPath), Path.GetFileNameWithoutExtension(strPDFPath) + ".xlsx"); ;
                //workbook.Save();
                // Make Excel visible to the user
                excelApp.Visible = true;

                // Release the Excel application object
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);


            }
            catch
            {
                workbook.Close();
                excelApp.Quit();
            }
        }

        private static int GetKeyFromValue(this Dictionary<int, double> dict, double value)
        {
            foreach (KeyValuePair<int, double> pair in dict)
            {
                if (pair.Value == value)
                {
                    return pair.Key;
                }
            }

            // If the value is not found, return null or raise an exception
            return -1;
        }
        private static string getValue(this Worksheet wk, int row, int column)
        {
            Range cellRange = wk.Range[wk.Cells[row, column], wk.Cells[row, column]];
            return Convert.ToString(cellRange.Value);
        }

        private static string getValueDT(this Worksheet wk, int row, int column)
        {
            try
            {
                Range cellRange = wk.Range[wk.Cells[row, column], wk.Cells[row, column]];
                if (cellRange.Value == null)
                    return default(string);
                try
                {
                    return Convert.ToDateTime(cellRange.Value).ToString("dd/MM/yyyy");
                }
                catch { }
                string dateString = Convert.ToString(cellRange.Value);
                string[] formats = { "dd-MMM-yyyy", "dd MMMM yyyy", "d/M/yyyy", "dd/M/yyyy", "d/MM/yyyy", "dd/MM/yyyy", "dd-MM-yyyy", "d-M-yyyy", "dd MMM", "yyyy/MM/dd", "dd MMM yyyy", "MMMM dd, yyyy", "MMM dd, yyyy", "yyyy MMMM dd", "yyyy-MMM-dd", "yyyy MMM dd", "MMM yyyy", "MM/dd/yyyy", "yyyy-MM-dd", "MM-dd-yyyy", "yyyy.MM.dd",
                "dd-MMM-yy", "dd MMMM yy", "d/M/yy", "dd/M/yy", "d/MM/yy", "dd/MM/yy", "dd-MM-yy", "d-M-yy", "dd MMM", "yy/MM/dd", "dd MMM yy", "MMMM dd, yy", "MMM dd, yy", "yy MMMM dd", "yy-MMM-dd", "yy MMM dd", "MMM yy", "MM/dd/yy", "yy-MM-dd", "MM-dd-yy", "yy.MM.dd" }; // add more formats as needed

                foreach (string format in formats)
                {
                    if (DateTime.TryParseExact(dateString, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
                    {
                        return date.ToString("dd/MM/yyyy");
                    }
                }
            }
            catch {
                return default(string);
            }
            return default(string);
        }

    }

    internal class Book
    {
        public Book(string bw,string bd, string bdesc)
        {
            this.BookDeposit = bd;
            this.BookWithdrawal = bw;
            this.BookDescription = bdesc;

        }

        public string BookDescription { get; set; }
        public string BookWithdrawal { get; set; }
        public string BookDeposit { get; set; }
    }
}
