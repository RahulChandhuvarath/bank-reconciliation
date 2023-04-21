using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Ghostscript.NET.Rasterizer;
using System.Drawing;
using Ghostscript.NET;
using System.IO;
using System.Drawing.Imaging;

namespace Reconciliation
{
    internal class PdfToImage
    {
        public static List<string> ConvertToImage(string pdfPath)
        {
            List<string> paths = new List<string>();
            // Load the PDF file into a GhostscriptRasterizer object
            using (var rasterizer = new GhostscriptRasterizer())
            {
                rasterizer.Open(pdfPath);

                // Loop through each page in the PDF file
                for (int pageNumber = 1; pageNumber <= rasterizer.PageCount; pageNumber++)
                {
                    // Set the resolution for the output image
                    var settings = new GhostscriptImageDevice();
                    settings.GraphicsAlphaBits = GhostscriptImageDeviceAlphaBits.V_4;
                    settings.TextAlphaBits = GhostscriptImageDeviceAlphaBits.V_4;
                    settings.ResolutionXY = new GhostscriptImageDeviceResolution(600, 600);


                    // Render the current page of the PDF file to a PNG image
                    using (var image = rasterizer.GetPage(300, pageNumber))
                    {
                        string fullPath = Path.Combine(Path.GetTempPath(), Path.GetFileNameWithoutExtension(pdfPath) + pageNumber.ToString() + ".png");
                        //string fullPath2 = Path.Combine(Path.GetDirectoryName(pdfPath), Path.GetFileNameWithoutExtension(pdfPath) + pageNumber.ToString() + "_X.png");
                        paths.Add(fullPath);
                        image.Save(fullPath);

                        //MakeGrayscale(fullPath, fullPath2);
                    }


                }

            }
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
