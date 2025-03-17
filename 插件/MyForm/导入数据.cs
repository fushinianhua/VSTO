using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using 插件.Properties;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

namespace 插件.MyForm
{
    public partial class 导入数据 : Form
    {
        public 导入数据()
        {
            InitializeComponent();
        }

        private string 数据导入地址 = Settings.Default.数据导入地址;
        private System.Data.DataTable dataTable = new System.Data.DataTable(); // 用于存储所有选中的列数据
        private void 导入数据_FormClosed(object sender, FormClosedEventArgs e)
        {


            Globals.ThisAddIn.导入form = null;
        }

        private void 导入数据_Load(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(数据导入地址))
                {
                    PathText.Text = 数据导入地址 = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                }
                currentSheet = Globals.ThisAddIn.Application.ActiveSheet;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载初始路径时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog fileDialog = new OpenFileDialog
                {
                    Title = "请选择需要导入的文件",
                    Filter = "Excel 文件 (*.xlsx;*.xls)|*.xlsx;*.xls|所有文件 (*.*)|*.*",
                    Multiselect = false,
                    InitialDirectory = 数据导入地址
                };

                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    数据导入地址 = fileDialog.FileName;

                    // 使用 NPOI 读取 Excel 文件
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

                        comboBox1.Items.Clear();
                        for (int i = 0; i < workbook.NumberOfSheets; i++)
                        {
                            comboBox1.Items.Add(workbook.GetSheetName(i));
                        }
                        comboBox1.SelectedIndex = comboBox1.Items.Count > 0 ? 0 : -1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载文件时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                CheckList.Items.Clear();
                if (comboBox1.SelectedIndex >= 0)
                {
                    // 使用 NPOI 读取选中的 Sheet
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

                        ISheet sheet = workbook.GetSheet(comboBox1.SelectedItem.ToString());
                        IRow headerRow = sheet.GetRow(0); // 假设表头在第一行

                        if (headerRow != null)
                        {
                            for (int i = 0; i < headerRow.LastCellNum; i++)
                            {
                                CheckList.Items.Add(headerRow.GetCell(i)?.ToString());
                                CheckList.SetItemChecked(i, true);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载表头时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        Microsoft.Office.Interop.Excel.Worksheet currentSheet;
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (comboBox1.SelectedIndex < 0)
                {
                    MessageBox.Show("请先选择一个 Sheet。", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // 清空 DataTable
                dataTable.Clear();
                dataTable.Columns.Clear();

                // 使用 NPOI 读取选中的 Sheet
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

                    ISheet sheet = workbook.GetSheet(comboBox1.SelectedItem.ToString());

                    // 读取表头
                    IRow headerRow = sheet.GetRow(0);
                    for (int i = 0; i < headerRow.LastCellNum; i++)
                    {
                        if (CheckList.GetItemChecked(i))
                        {
                            dataTable.Columns.Add(headerRow.GetCell(i)?.ToString());
                        }
                    }

                    // 读取数据
                    for (int row = 1; row <= sheet.LastRowNum; row++)
                    {
                        IRow currentRow = sheet.GetRow(row);
                        if (currentRow != null)
                        {
                            DataRow dataRow = dataTable.NewRow();
                            int colIndex = 0;
                            for (int col = 0; col < headerRow.LastCellNum; col++)
                            {
                                if (CheckList.GetItemChecked(col))
                                {
                                    dataRow[colIndex] = currentRow.GetCell(col)?.ToString() ?? "0";
                                    colIndex++;
                                }
                            }
                            dataTable.Rows.Add(dataRow);
                        }
                    }
                }

                // 将数据写入 currentSheet

                int lastRow = currentSheet.Cells[currentSheet.Rows.Count, 1].End[XlDirection.xlUp].Row + 1;

                for (int row = 0; row < dataTable.Rows.Count; row++)
                {
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        currentSheet.Cells[lastRow + row, col + 1].Value2 = dataTable.Rows[row][col];
                    }
                }

                MessageBox.Show("数据导入成功！", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导入数据时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (checkBox1.Checked)
                {
                    for (int i = 1; i <= CheckList.Items.Count; i++)
                    {
                        CheckList.SetItemChecked(i - 1, true);
                        checkBox1.Text = "全部取消";
                    }

                }
                else
                {
                    for (int i = 1; i <= CheckList.Items.Count; i++)
                    {
                        CheckList.SetItemChecked(i - 1, false);
                        checkBox1.Text = "全部选中";
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}