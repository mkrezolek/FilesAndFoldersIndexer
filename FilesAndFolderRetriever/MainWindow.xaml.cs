using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
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
using System.Windows.Threading;
using WinForms = System.Windows.Forms;


namespace FilesAndFolderRetriever
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string folderPath;

        
        public MainWindow()
        {
            InitializeComponent();
            
        }

        public void BtnFolderBrowse_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new WinForms.FolderBrowserDialog();
            WinForms.DialogResult result =  dialog.ShowDialog();
            folderPath = dialog.SelectedPath;
            txtPath.Text = folderPath;             
        }

        
        private void BtnSaveAs_Click(object sender, RoutedEventArgs e)
        {
            WinForms.SaveFileDialog saveAsDialog  = new WinForms.SaveFileDialog();
            saveAsDialog.Filter = "XLS files (*.xls)|*.xls";
            WinForms.DialogResult saveAsResult = saveAsDialog.ShowDialog();
            txtSaveToPath.Text = saveAsDialog.FileName;
        }

       

        private void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            
            var watch = new System.Diagnostics.Stopwatch();

            watch.Start();

            //Create DIrectorySearch object from given path
            DirectorySearch directorySearch = new DirectorySearch(txtPath.Text);
            //Create list of files returned by traversing the directory
            IEnumerable<System.IO.FileInfo> files = directorySearch.TraverseDirectory();
            Thread.Sleep(100);
            prgBar.Maximum = files.Count();

            //Create writeToExcel object from given path
            WriteToExcel writeToExcel = new WriteToExcel(txtSaveToPath.Text);
            //Setup Excel file prior to data input
            writeToExcel.setupExcel();
            //Iterate through the list of files and add data to the Excel file
            writeToExcel.addData(files);
            //Save the file and tidy up
            writeToExcel.saveExcelFile();


            watch.Stop();

            MessageBox.Show($"Excel file created in {watch.ElapsedMilliseconds} ms");
            
        }


    }
}
