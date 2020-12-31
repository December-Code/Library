using System;
using ExcelDataReader;
using System.Text;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
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
        DataSet result;
        private void Upload_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog();
            string filePath = null;
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                openFile.RestoreDirectory = false;
                openFile.Title = "選擇你要的檔案"; //開窗標題
                filePath = openFile.FileName;
                FileName.Text = filePath;
            }

            if (filePath != null)
            {
                if (File.Exists(filePath))
                {
                    var extension = Path.GetExtension(filePath).ToLower();
                    using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        IExcelDataReader reader = null;
                        if (extension == ".xls")
                        {

                            status_content.Text = "XLS格式";
                            reader = ExcelReaderFactory.CreateBinaryReader(stream, new ExcelReaderConfiguration()
                            {
                                FallbackEncoding = Encoding.GetEncoding("big5")
                            });
                        }
                        else if (extension == ".xlsx")
                        {
                            status_content.Text = "XLSX格式";
                            reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                        }
                        else if (extension == ".csv")
                        {
                            status_content.Text = " CSV格式";
                            reader = ExcelReaderFactory.CreateCsvReader(stream, new ExcelReaderConfiguration()
                            {
                                FallbackEncoding = Encoding.GetEncoding("big5")
                            });
                        }
                        /////////////////////////////////////////讀取檔案結束///////////////////////////////////
                        if (reader == null)
                        {
                            MessageBox.Show("錯誤檔案格式：" + extension);
                            status_content.Text = "錯誤檔案格式";
                        }
                        else
                        {
                            Execel_preview.DataSource = "";
                            using (reader)
                            {
                                result = reader.AsDataSet(new ExcelDataSetConfiguration()
                                {
                                    UseColumnDataType = false,
                                    ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                                    {
                                        //設定讀取資料時是否忽略標題
                                        UseHeaderRow = false
                                    }
                                });
                                //把 DataSet 顯示出來
                                DataTable tableOutput = result.Tables[0];
                                Execel_preview.DataSource = tableOutput;
                                uploaded = true;
                            }
                        }
                    }//using FileStream 內
                }
                else
                {
                    MessageBox.Show("檔案不存在");
                }
            }

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
            FolderBrowserDialog Save = new FolderBrowserDialog();
            if (Save.ShowDialog() == DialogResult.OK)
            {                
                //Creae an Excel application instance
                Excel.Application excelApp = new Excel.Application();

                //Create an Excel workbook instance and open it from the predefined location
                Excel.Workbook excelWorkBook = excelApp.Workbooks.Add(1);

                foreach (DataTable table in ds.Tables)
                {
                    //Add a new worksheet to workbook with the Datatable name
                    Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();

                    excelWorkSheet.Name = table.TableName;

                    for (int i = 1; i < table.Columns.Count + 1; i++)
                    {
                        excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
                    }

                    for (int j = 0; j < table.Rows.Count; j++)
                    {
                        for (int k = 0; k < table.Columns.Count; k++)
                        {
                            excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
                        }
                    }
                }
                excelWorkBook.SaveAs(@Save+"Test.xlsx");
                MessageBox.Show("完成囉!!");
                excelWorkBook.Close();
                excelApp.Quit();
            }
        }
    }
}

