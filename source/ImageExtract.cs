using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tesseract;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;
using ImageProcessor.Imaging.Filters.Photo;
using Microsoft.Office.Interop.Excel;
using ImageProcessor;
using System.Reflection;
using System.Resources;
using System.Globalization;

namespace Reconciliation
{
    internal class ImageExtract
    {
        public static List<AccountRow> TextFromImage(List<string> imagePaths)
        {


            StringBuilder streamText = new StringBuilder();

            foreach (string imageFullPath in imagePaths)
            {
                // Load the image using ImageProcessor
                using (var imageFactory = new ImageFactory(preserveExifData: true))
                {
                    imageFactory.Load(imageFullPath);

                    // Apply contrast filter to the image to improve its quality
                    imageFactory.Contrast(25);

                    // Convert the image to grayscale to reduce noise and enhance text
                    // Convert the image to grayscale
                    imageFactory.Filter(MatrixFilters.GreyScale);


                    // Apply Gaussian blur to the image to further reduce noise and smooth edges
                    imageFactory.GaussianBlur(5);

                    string fullPath2 = Path.Combine(Path.GetDirectoryName(imageFullPath), Path.GetFileNameWithoutExtension(imageFullPath)  + "_Processed.png");
                    // Save the processed image to a new file for debug purposes
                    imageFactory.Save(fullPath2);

                    string language = "eng";
                    string tessdatapath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
                    //var traineddataPath = Path.Combine(tessdatapath, "eng.traineddata");
                    //using (var stream = GetEmbeddedResource("tessdata." + language + ".traineddata"))
                    //{
                    //    using (var fileStream = new FileStream(tessdatapath, FileMode.Create))
                    //    {
                    //        stream.CopyTo(fileStream);
                    //    }
                    //}
                    var engine = new TesseractEngine(tessdatapath, language, EngineMode.Default);

                    // Set up the Tesseract OCR engine with custom options
                    //var engine = new TesseractEngine(@"C:\Projects\BankReconciliation\packages\Tesseract.Data.English.4.0.0\build\tessdata", "eng", EngineMode.Default);

                    engine.DefaultPageSegMode = PageSegMode.SingleBlock;

                    engine.SetVariable("textord_tabfind_find_hlines", "0"); // Disable finding horizontal lines in the text
                    engine.SetVariable("textord_tabfind_find_vlines", "0"); // Disable finding vertical lines in the text
                 


                    var pix = Pix.LoadFromFile(fullPath2);

                    using (var page = engine.Process(pix))
                    {
                        var text = page.GetText();
                        streamText.Append(text);

                    }
                }
            }

            // Extract the text from the OCR result
            string tableText = streamText.ToString();

            List<AccountRow> accounts = new List<AccountRow>();
            // Split the text into rows
            string[] rows = tableText.Split('\n');
            foreach (var row in rows)
            {
                string line = row;
                string date = "";
                double withdrawal = 0;
                double deposit = 0;
                double balance = 0;
                List<double> lstAmounts = new List<double>();


                //string datePattern = @"\b\d{1,2}[-/]\d{1,2}[-/]\d{2,4}\b|\b\d{1,2}\s+\w{3}\s+\d{2,4}\b|\b\w{3}\s+\d{1,2},\s+\d{4}\b";
                string datePattern = @"\b\d{1,2}[-/]\d{1,2}[-/](\d{2,4})?\b|\b\d{1,2}\s+\w{3}\s+(\d{2,4})?\b|\b\w{3}\s+\d{1,2},\s+(\d{4})?\b";

                MatchCollection dataMatches = Regex.Matches(line, datePattern);
                if (dataMatches.Count == 0)
                    continue;

                date = ToDT(dataMatches[0].Value);
                foreach (Match item in dataMatches)
                {
                    line = line.Replace(item.Value,"");
                }
                string numPattern = @"\b\d{1,3}(,\d{3})*(\.\d+)\b|\b\d+(\.\d+)\b";
                //string numPattern = @"\b\d{1,3}(,\d{3})*(\.\d+)\b";//mandatory decimal optional commas
                //string numPattern = @"\b\d{1,3}(,\d{3})*\.\d+\b";  //mandatory decimal
                //string pattern = @"\b\d{1,3}(,\d{3})*(\.\d+)?\b"; //optional decimal
                Regex regex = new Regex(numPattern);

                
                MatchCollection ammountMatches = regex.Matches(line);
                //lstAmounts = ammountMatches.Cast<Match>().Select(m => m.Value).ToList();
                foreach (Match item in ammountMatches)
                {
                    lstAmounts.Add(Convert.ToDouble(item.Value.Replace(",","")));
                }
                if (lstAmounts.Count() < 2)
                    continue;
                accounts.Add(new AccountRow { Date = date, Withdrawal = withdrawal, Deposit=deposit, Balance = balance, ListAmount =lstAmounts });
                
            }

            List<AccountRow> accounttemp1 = accounts.ConvertAll(x => new AccountRow(x));
           
            List<AccountRow> accounttemp2= accounts.ConvertAll(x => new AccountRow(x));
            accounttemp2.Reverse();

            GetAccountList(ref accounttemp1);
            GetAccountList(ref accounttemp2);
            accounts = new List<AccountRow>(accounttemp1);
            if(accounttemp2.Count > accounttemp1.Count)
                accounts = new List<AccountRow>(accounttemp2);

            return accounts;

         
           
        }

        static Stream GetEmbeddedResource(string resourceName)
        {
            var assembly = Assembly.GetExecutingAssembly();
            var resourceNames = assembly.GetManifestResourceNames();
            var fullName = resourceNames.FirstOrDefault(name => name.EndsWith(resourceName));
            if (fullName == null)
            {
                throw new Exception("Embedded resource not found: " + resourceName);
            }
          
            return assembly.GetManifestResourceStream(fullName);
        }

        private static string ToDT(string dateString)
        {

            string[] formats = { "dd-MMM-yyyy", "dd MMMM yyyy", "d/M/yyyy", "dd/M/yyyy", "d/MM/yyyy", "dd/MM/yyyy", "dd-MM-yyyy", "d-M-yyyy", "dd MMM", "yyyy/MM/dd", "dd MMM yyyy", "MMMM dd, yyyy", "MMM dd, yyyy", "yyyy MMMM dd", "yyyy-MMM-dd", "yyyy MMM dd", "MMM yyyy", "MM/dd/yyyy", "yyyy-MM-dd", "MM-dd-yyyy", "yyyy.MM.dd" }; // add more formats as needed

            foreach (string format in formats)
            {
                if (DateTime.TryParseExact(dateString, format, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
                {
                    return date.Date.ToString("dd/MM/yyyy");
                }
            }

            return dateString;

        }

        private static void GetAccountList(ref List<AccountRow> accounts)
        {

            double previousBalance = -1;
            List<List<double>> lstAmountInitial = new List<List<double>>();
            for (int i = 0; i < accounts.Count; i++)
            {
                if (i != accounts.Count - 1)
                {
                    var acc = GetAccounts(accounts[i].ListAmount, accounts[i + 1].ListAmount);
                    if (acc.Item4 != "")
                    {
                        lstAmountInitial.Add(new List<double>(accounts[i].ListAmount.Intersect(new List<double>() { acc.Item1, acc.Item2 })));
                    }
                    if (acc.Item4 == "add")
                    {
                        if (previousBalance == -1)
                        {

                        }
                        else
                        {
                            if (acc.Item1 == previousBalance)
                            {
                                accounts[i].Balance = acc.Item1;
                                accounts[i].Withdrawal = acc.Item2;
                            }
                            else
                            {
                                accounts[i].Balance = acc.Item2;
                                accounts[i].Withdrawal = acc.Item1;
                            }
                        }
                        previousBalance = acc.Item3;
                    }
                    else if (acc.Item4 == "sub")
                    {
                        accounts[i].Balance = Math.Max(acc.Item1, acc.Item2);
                        accounts[i].Deposit = Math.Min(acc.Item1, acc.Item2);

                        previousBalance = acc.Item3;
                    }
                    else
                    {
                        accounts.RemoveAt(i);
                        i--;
                    }
                }
                else
                {
                    if (previousBalance != -1)
                    {
                        accounts[i].Balance = previousBalance;

                        if(accounts[i].ListAmount.Count ==2)
                        {
                            List<double> lstTemp = new List<double>(accounts[i].ListAmount);
                            lstTemp.Remove(previousBalance);
                            if (lstTemp[0] > previousBalance)
                                accounts[i].Withdrawal = lstTemp[0];
                            else
                                accounts[i].Deposit = lstTemp[0];
                        }
                    }
                }
            }
            if (lstAmountInitial.Count >= 2 && accounts[0].Balance == 0)
            {
                accounts[0].Balance = lstAmountInitial[0][lstAmountInitial[1].IndexOf(accounts[1].Balance)];
                accounts[0].Withdrawal = lstAmountInitial[0].Except(new List<double>() { accounts[0].Balance }).ToList()[0];
            }
          
        }

        public static string GreyScale(string imageFile)
        {
            Bitmap image = new Bitmap(imageFile);
            // Convert the image to grayscale
            var grayscaleBitmap = new Bitmap(image.Width, image.Height);
            for (int x = 0; x < grayscaleBitmap.Width; x++)
            {
                for (int y = 0; y < grayscaleBitmap.Height; y++)
                {
                    var color = image.GetPixel(x, y);
                    var gray = (int)(0.299 * color.R + 0.587 * color.G + 0.114 * color.B);
                    grayscaleBitmap.SetPixel(x, y, Color.FromArgb(gray, gray, gray));
                }
            }

            // Threshold the image to black and white
            var thresholdBitmap = new Bitmap(grayscaleBitmap.Width, grayscaleBitmap.Height);
            for (int x = 0; x < thresholdBitmap.Width; x++)
            {
                for (int y = 0; y < thresholdBitmap.Height; y++)
                {
                    var color = grayscaleBitmap.GetPixel(x, y);
                    var gray = (int)(0.299 * color.R + 0.587 * color.G + 0.114 * color.B);
                    if (gray < 128)
                    {
                        thresholdBitmap.SetPixel(x, y, Color.Black);
                    }
                    else
                    {
                        thresholdBitmap.SetPixel(x, y, Color.White);
                    }
                }
            }

            // Save the thresholded image to a temporary file
            string outFile = Path.Combine(Path.GetTempPath(), Path.GetFileName(imageFile));
            thresholdBitmap.Save(outFile);


            return outFile;
        }


        private static (double, double, double, string) GetAccounts(List<double> listA, List<double> listB)
        {
           
            foreach (double a in listA)
            {
                foreach (double b in listA)
                {
                    if (a != b)
                    {
                        double sum = a + b;
                        double difference = Math.Abs( a - b);

                        bool containsNumber = listB.Any(d => Math.Truncate(d) == Math.Truncate(sum));

                        if (containsNumber)
                        {
                            double matchingNumber = listB.First(d => Math.Truncate(d) == Math.Truncate(sum));
                            return (a, b, matchingNumber, "add");
                        }
                        containsNumber = listB.Any(d => Math.Truncate(d) == Math.Truncate(difference));
                        if (containsNumber)
                        {
                            double matchingNumber = listB.First(d => Math.Truncate(d) == Math.Truncate(difference));
                            return (a, b, matchingNumber, "sub");
                        }
                    }
                }
            }
            return (default(double), default(double), default(double), "");
        }
    }

    class AccountRow
    {
        public AccountRow(AccountRow x)
        {
            this.Date = x.Date;
            this.Balance = x.Balance;
            this.Withdrawal = x.Withdrawal;
            this.Deposit = x.Deposit;
            this.ListAmount = new List<double>( x.ListAmount);
        }
        public AccountRow()
        {

        }
        public string Date { get; set; }
        public double Withdrawal { get; set; }
        public double Deposit { get; set; }

        public double BookWithdrawal { get; set; }
        public double BookDeposit { get; set; }

        public double Balance { get; set; }
        public List<double> ListAmount { get; set; }
    }
}
