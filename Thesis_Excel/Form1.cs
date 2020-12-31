using System;
using System.Text;
using System.Data;
using System.Windows.Forms;
using System.IO;

namespace Thesis_Excel
{
    public partial class main : Form
    {
        public main()
        {
            InitializeComponent();
        }
        bool uploaded = false;
        private void Upload_Click(object sender, EventArgs e)
        {
            //OpenFileDialog openFile = new OpenFileDialog();
            //string filePath = null;
            //if (openFile.ShowDialog() == DialogResult.OK)
            //{
            //    openFile.RestoreDirectory = false;
            //    openFile.Title = "選擇你要的檔案"; //開窗標題
            //    filePath = openFile.FileName;
            //    FileName.Text = filePath;
            //}     

        }

        private void Convert_Click(object sender, EventArgs e)
        {
            if (uploaded)
            {
                DataSet DatasetExport = new DataSet();
                DataTable tableExport = result.Tables[0];
                for (int row = 0; row < tableExport.Rows.Count; row++)
                {
                    for (var col = 0; col < tableExport.Columns.Count; col++)
                    {
                        string data = tableExport.Rows[row][col].ToString();
                    }
                }
                DatasetExport.Tables.Add(tableExport.Copy());

                ExportDataSetToExcel(DatasetExport);
            }
            else
            {
                MessageBox.Show("請先匯入檔案");
            }
        }
        private static void ExportDataSetToExcel(DataSet ds)
        {
            FileInfo fileInfo = new FileInfo(pathName);

            using (ExcelPackage packge = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = packge.Workbook.Worksheets[sheetName];

                for (int i = 0; i < newInfoArray.Length; ++i)
                {
                    worksheet.Cells[selectIndex + 2, i + 1].Value = newInfoArray[i];
                }

                packge.Save();
            }
        }
        public static void WriteToExcel(string[] newInfoArray)
        {
            FileInfo fileInfo = new FileInfo(pathName);

            using (ExcelPackage packge = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = packge.Workbook.Worksheets[sheetName];

                for (int i = 0; i < newInfoArray.Length; ++i)
                {
                    worksheet.Cells[selectIndex + 2, i + 1].Value = newInfoArray[i];
                }

                packge.Save();
            }
        }
        //private static void ExportDataSetToExcel(DataSet ds)
        //{
        //    FolderBrowserDialog Save = new FolderBrowserDialog();
        //    if (Save.ShowDialog() == DialogResult.OK)
        //    {                
        //        //Creae an Excel application instance
        //        Excel.Application excelApp = new Excel.Application();

        //        //Create an Excel workbook instance and open it from the predefined location
        //        Excel.Workbook excelWorkBook = excelApp.Workbooks.Add(1);

        //        foreach (DataTable table in ds.Tables)
        //        {
        //            //Add a new worksheet to workbook with the Datatable name
        //            Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();

        //            excelWorkSheet.Name = table.TableName;

        //            for (int i = 1; i < table.Columns.Count + 1; i++)
        //            {
        //                excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
        //            }

        //            for (int j = 0; j < table.Rows.Count; j++)
        //            {
        //                for (int k = 0; k < table.Columns.Count; k++)
        //                {
        //                    excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
        //                }
        //            }
        //        }
        //        excelWorkBook.SaveAs(@Save+"Test.xlsx");
        //        MessageBox.Show("完成囉!!");
        //        excelWorkBook.Close();
        //        excelApp.Quit();
        //    }
        //}
    }
}

