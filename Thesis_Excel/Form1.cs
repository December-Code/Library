using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using ExcelDataReader;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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


            DataSet result;
            DataRowCollection dataRow;
            DataColumnCollection dataColumn;

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
                            var table = result.Tables[0];
                            for (int row = 0; row < table.Rows.Count; row++)
                            {
                                for (var col = 0; col < table.Columns.Count; col++)
                                {
                                    string data = table.Rows[row][col].ToString();                            
                                }
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
    }
}

