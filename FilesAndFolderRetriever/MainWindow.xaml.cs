using System;
using System.Collections.Generic;
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
using WinForms = System.Windows.Forms;


namespace FilesAndFolderRetriever
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string folderPath = "";

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
            DirectorySearch directorySearch = new DirectorySearch(folderPath);
            var files = directorySearch.TraverseDirectory();
            WriteToExcel writeToExcel = new WriteToExcel(txtSaveToPath.Text);
            writeToExcel.setupExcel();
            writeToExcel.addData(files);
            writeToExcel.saveExcelFile();
            watch.Stop();
            
            MessageBox.Show($"Excel file created in {watch.ElapsedMilliseconds} ms");

        }


    }
}
