using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;


namespace FilesAndFolderRetriever
{
    class WriteToExcel
    {
        public string filePath;
        public Excel.Workbook xlWorkbook;
        public Excel.Worksheet xlWorksheet;
        private Excel.Application xlApp = new Excel.Application();
        public object misValue = System.Reflection.Missing.Value;

        public WriteToExcel(string path)
        {
            this.filePath = path;            
        }

        public void setupExcel()
        {            
            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!");
            }
            else
            {                              
                
                xlWorkbook = xlApp.Workbooks.Add(misValue);
                xlWorksheet = (Excel.Worksheet)xlWorkbook.Worksheets.get_Item(1);
            }                       
                       
        }

        public void addData(IEnumerable<FileInfo> files)
        {
            int row = 1;
                        
            xlWorksheet.Cells[1, 1] = "Full Name";
            xlWorksheet.Cells[1, 2] = "File Name";
            xlWorksheet.Cells[1, 3] = "Size";
            xlWorksheet.Cells[1, 4] = "Directory Name";
            xlWorksheet.Cells[1, 5] = "Last Access Time";
            xlWorksheet.Cells[1, 6] = "Last Write Time";
            xlWorksheet.Cells[1, 7] = "Creation Time";


            row = 2;
            
            Parallel.ForEach(files, file =>
            {
                xlWorksheet.Cells[row, 1] = file.FullName;
                xlWorksheet.Cells[row, 2] = file.Name;
                xlWorksheet.Cells[row, 3] = file.Length;
                xlWorksheet.Cells[row, 4] = file.DirectoryName;
                xlWorksheet.Cells[row, 5] = file.LastAccessTime;
                xlWorksheet.Cells[row, 6] = file.LastWriteTime;
                xlWorksheet.Cells[row, 7] = file.CreationTime;

                row++;
            });
            
        }       

        public void saveExcelFile()
        {
            xlWorkbook.SaveAs(filePath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue);
            xlWorkbook.Close(true, misValue, misValue);
            xlApp.Quit();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

        }

    }
}
