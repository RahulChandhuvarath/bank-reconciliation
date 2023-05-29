using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;
using Spire.Pdf;
using Spire.Pdf.Graphics;
using System.Diagnostics;



namespace Reconciliation
{
    internal class PdfToImage
    {
        public static List<string> ConvertToImage(string pdfPath)
        {
            List<string> paths = new List<string>();
         

            PdfDocument pdfDocument = new PdfDocument();
            pdfDocument.LoadFromFile(pdfPath);

            for (int i = 0; i < pdfDocument.Pages.Count; i++)
            {

                string fullPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(pdfPath) +"_"+ (i+1).ToString() + ".png");
                paths.Add(fullPath);
            }

            pdfDocument.Close();

            // Create a new process instance
            Process process = new Process();
          
            // Configure the process start info
            process.StartInfo.FileName = Path.Combine(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "Pdf2Image.exe");   // Specify the path to your EXE file
            process.StartInfo.CreateNoWindow = true;  // Set to true to hide the command prompt window
            process.StartInfo.UseShellExecute = false;  // Set to false to redirect input/output

            process.StartInfo.Arguments = $"\"{pdfPath}\"";

            // Start the process
            process.Start();

            return paths;

        }



        public static void MakeGrayscale(string inputFilePath, string outputFilePath)
        {
            // Load the input image
            Bitmap inputImage = new Bitmap(inputFilePath);

            // Create a grayscale color matrix
            ColorMatrix grayscaleMatrix = new ColorMatrix(
                new float[][] {
            new float[] {0.299f, 0.299f, 0.299f, 0, 0},
            new float[] {0.587f, 0.587f, 0.587f, 0, 0},
            new float[] {0.114f, 0.114f, 0.114f, 0, 0},
            new float[] {0, 0, 0, 1, 0},
            new float[] {0, 0, 0, 0, 1}
                });

            // Create an ImageAttributes object with the grayscale matrix
            ImageAttributes attributes = new ImageAttributes();
            attributes.SetColorMatrix(grayscaleMatrix);

            // Create a new Bitmap with the same dimensions as the input image
            Bitmap outputImage = new Bitmap(inputImage.Width, inputImage.Height);

            // Draw the input image onto the output image using the grayscale color matrix
            using (Graphics graphics = Graphics.FromImage(outputImage))
            {
                graphics.DrawImage(inputImage, new Rectangle(0, 0, inputImage.Width, inputImage.Height),
                    0, 0, inputImage.Width, inputImage.Height, GraphicsUnit.Pixel, attributes);
            }

            // Save the output image as a new file
            outputImage.Save(outputFilePath, ImageFormat.Png);
        }

    }
}
