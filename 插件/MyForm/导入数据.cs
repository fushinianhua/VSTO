using Microsoft.Office.Interop.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using 插件.Properties;

namespace 插件.MyForm
{
    public partial class 导入数据 : Form
    {
        public 导入数据()
        {
            InitializeComponent();
        }
        private string 数据导入地址 = Settings.Default.数据导入地址;
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog fileDialog = new OpenFileDialog
                {
                    Title = "请选择需要导入的文件",
                    Filter = "Excel 文件 (*.xlsx;*.xls)|*.xlsx;*.xls|文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*",
                    Multiselect = false,
                    InitialDirectory = 数据导入地址
                };
                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    数据导入地址 = fileDialog.FileName;
                    using (FileStream fs = new FileStream(数据导入地址, FileMode.Open, FileAccess.Read))
                    {
                        IWorkbook workbook;
                        if (Path.GetExtension(数据导入地址).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                        {
                            workbook = new XSSFWorkbook(fs);
                        }
                        else
                        {
                            workbook = new HSSFWorkbook(fs);
                        }

                        for (int i = 0; i < workbook.NumberOfSheets; i++)
                        {
                            ISheet sheet = workbook.GetSheetAt(i);
                            comboBox1.Items.Add(sheet.SheetName);
                        }
                        comboBox1.SelectedIndex = comboBox1.Items.Count > 0 ? 0 : -1;

                        // 这里可以进一步读取指定工作表的数据
                        if (comboBox1.SelectedIndex >= 0)
                        {
                            string selectedSheetName = comboBox1.SelectedItem.ToString();
                            ISheet selectedSheet = workbook.GetSheet(selectedSheetName);
                            int rowCount = selectedSheet.LastRowNum + 1;
                            for (int row = 0; row < rowCount; row++)
                            {
                                IRow currentRow = selectedSheet.GetRow(row);
                                if (currentRow != null)
                                {
                                    int colCount = currentRow.LastCellNum;
                                    for (int col = 0; col < colCount; col++)
                                    {
                                        ICell cell = currentRow.GetCell(col);
                                        var cellValue = cell?.ToString();
                                        // 处理单元格数据
                                    }
                                }
                            }
                        }
                    }
                    comboBox1.SelectedIndex = comboBox1.Items.Count > 0 ? 0 : -1;
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                string fileExtension = Path.GetExtension(数据导入地址);
                switch (fileExtension)
                {
                    case ".xlsx":
                    case ".xls":
                        // 处理 Excel 文件
                        ReadExcelFile(数据导入地址);
                        break;
                    case ".txt":
                        // 处理文本文件
                        ReadTextFile(数据导入地址);
                        break;
                    default:
                        MessageBox.Show("不支持的文件类型。");
                        break;
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
        Workbook workbook; Worksheet worksheet;
        private void ReadExcelFile(string filePath)
        {
            try
            {

                Range range = worksheet.UsedRange;





                workbook.Close(false);

                // 释放 COM 对象
                System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"读取 Excel 文件时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ReadTextFile(string filePath)
        {
            try
            {
                string[] lines = File.ReadAllLines(filePath);
                // 这里可以对读取的文本行进行进一步处理，例如显示在 TextBox 中
                MessageBox.Show($"成功读取文本文件，包含 {lines.Length} 行。");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"读取文本文件时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void 导入数据_FormClosed(object sender, FormClosedEventArgs e)
        {
            Globals.ThisAddIn.导入form = null;
        }

        private void 导入数据_Load(object sender, EventArgs e)
        {
            try
            {
                if (数据导入地址 == "" || 数据导入地址 == null)
                {
                    PathText.Text = 数据导入地址 = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                }
            }
            catch (Exception)
            {

                throw;
            }

        }
        object[,] 表头数据 = null;
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Range r=null; Range rng=null;
            try
            {
                CheckList.Items.Clear();
            if (comboBox1.SelectedIndex >= 0)
                {
                    worksheet = workbook.Sheets[comboBox1.SelectedItem];
                    r = (Range)worksheet.Cells[1, worksheet.Columns.Count];//最后一列
                    int col = r.End[XlDirection.xlToLeft].Column;
                    rng = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, col]];
                    表头数据 = rng.Value2;
                    if (表头数据 != null)
                    {
                        for (int i = 1; i <= 表头数据.GetLength(1); i++)
                        {
                            CheckList.Items.Add(表头数据[1, i].ToString());
                        }
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                Marshal.ReleaseComObject(r);
                Marshal.ReleaseComObject(rng);

            }
        }
    }
}
