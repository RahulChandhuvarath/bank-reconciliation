using ImageProcessor.Common.Extensions;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Reconciliation
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, INotifyPropertyChanged
    {
        public MainWindow()
        {
            InitializeComponent();
            this.OverallProgress = 0;
        }

        private double overallProgress;
        public double OverallProgress
        {
            get
            {
                return this.overallProgress;
            }
            set
            {
                if (value != this.overallProgress)
                {
                    this.overallProgress = value;
                    NotifyPropertyChanged("OverallProgress");
                }
            }
        }

        protected void NotifyPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        public event PropertyChangedEventHandler PropertyChanged;
        private void Pdf_Browse(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();
            dialog.Filter = "PDF files (*.pdf)|*.pdf";
            if (dialog.ShowDialog() == true)
            {
                PDF_Path.Text = dialog.FileName;
            }
        }

        private void Excel_Browse(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();
            dialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            if (dialog.ShowDialog() == true)
            {
                Excel_Path.Text = dialog.FileName;
            }
        }

        private void Button_Execute(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(PDF_Path.Text) || string.IsNullOrEmpty(Excel_Path.Text))
            {
                MessageBox.Show("Please select both PDF and Excel files.");
                return;
            }

            try
            {

                Mouse.OverrideCursor = Cursors.Wait;
                var images = PdfToImage.ConvertToImage(PDF_Path.Text);
                var accountClass = ImageExtract.TextFromImage(images);
                ExcelOperation.ExcelReconciliation(accountClass, Excel_Path.Text);
            }
            catch { }
            finally
            {
                Mouse.OverrideCursor = Cursors.Arrow;
            }
        }
    }

}
