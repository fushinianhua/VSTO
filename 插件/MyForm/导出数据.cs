using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 插件.MyForm
{
    public partial class 导出数据 : Form
    {
        public 导出数据()
        {
            InitializeComponent();
            工作表 = StaticClass.ExcelApp.ActiveSheet;
        }
        private string 保存地址;
        private Worksheet 工作表 = null;
        List<string> 列名 = new List<string>();
        object[,] 单元格数据 = null;
        private void 导出数据_Load(object sender, EventArgs e)
        {
            Range r = null; Range c = null; Range 表头单元格 = null; Range 数据单元格 = null;
            try
            {
                if (工作表 != null)
                {
                    r = (Range)工作表.Cells[1, 工作表.Columns.Count];//最后一列
                    int col = r.End[XlDirection.xlToLeft].Column;
                    c = (Range)工作表.Cells[工作表.Rows.Count, 1];//最后一行
                    int row = c.End[XlDirection.xlUp].Row;
                    表头单元格 = 工作表.Range[工作表.Cells[1, 1], 工作表.Cells[1, col]];
                    数据单元格 = 工作表.Range[工作表.Cells[1, 1], 工作表.Cells[row, col]];
                    单元格数据 = 数据单元格.Value2;
                    for (int i = 1; i <= col; i++)
                    {
                        CheckList.Items.Add(单元格数据[1, i]);
                        CheckList.SetItemChecked(i - 1, true);
                    }
                    if (StaticClass.数据导出地址 == "")
                    {
                        保存地址 = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
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
                Marshal.ReleaseComObject(c);
                Marshal.ReleaseComObject(数据单元格);
                Marshal.ReleaseComObject(表头单元格);
            }
        }
        private void 导出数据_FormClosed(object sender, FormClosedEventArgs e)
        {
            Globals.ThisAddIn.导出form = null;
        }

        private void 文件导出_Click(object sender, EventArgs e)
        {
            if (CheckList.CheckedItems.Count > 0)
            {
                // 获取选中列的索引
                List<int> selectedColumnIndices = new List<int>();
                for (int i = 0; i < CheckList.Items.Count; i++)
                {
                    if (CheckList.GetItemChecked(i))
                    {
                        // 注意 Excel 索引从 1 开始，而 CheckList 索引从 0 开始
                        selectedColumnIndices.Add(i + 1);
                    }
                }
                int rowCount = 单元格数据.GetLength(0);
                int selectedColumnCount = selectedColumnIndices.Count;
                // 创建新的二维数组来存储选中列的数据
                object[,] newData = new object[rowCount, selectedColumnCount];
                // 复制选中列的数据到新数组
                for (int row = 0; row < rowCount; row++)
                {
                    for (int colIndex = 0; colIndex < selectedColumnCount; colIndex++)
                    {
                        int actualColumn = selectedColumnIndices[colIndex];
                        // 从原数据中取出对应行和选中列的数据放入新数组
                        newData[row, colIndex] = 单元格数据[row + 1, actualColumn];
                    }
                }
                // 以下是将新数据导出到新 Excel 工作表的示例
                Workbook newWorkbook = null;
                Worksheet newWorksheet = null;
                Range newRange = null;
                try
                {
                    newWorkbook = StaticClass.ExcelApp.Workbooks.Add();
                    string wbname = WBnameText.Text.Trim();
                    if (WBnameText.Text.Trim() == "")
                    {
                        wbname = $"{DateTime.Now.Year}-{DateTime.Now.Month}";
                    }
                    newWorksheet = newWorkbook.ActiveSheet;
                    if (WSnameText.Text.Trim() != "")
                    {
                        newWorksheet.Name = WSnameText.Text.Trim();
                    }
                    // 确定新工作表要写入数据的范围
                    newRange = newWorksheet.Range[
                        newWorksheet.Cells[1, 1],
                        newWorksheet.Cells[rowCount, selectedColumnCount]
                    ];
                    // 将新数据写入新工作表
                    newRange.Value2 = newData;
                    // 保存新工作簿
                    string str = Path.Combine(保存地址, wbname + ".xlsx");
                    newWorkbook.SaveAs(str);
                    newWorkbook.Close();
                    MessageBox.Show("数据导出成功！");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"数据导出失败：{ex.Message}");
                }
                finally
                {
                    // 释放资源
                    if (newRange != null)
                    {
                        Marshal.ReleaseComObject(newRange);
                        newRange = null;
                    }
                    if (newWorksheet != null)
                    {
                        Marshal.ReleaseComObject(newWorksheet);
                        newWorksheet = null;
                    }
                    if (newWorkbook != null)
                    {
                        Marshal.ReleaseComObject(newWorkbook);
                        newWorkbook = null;
                    }
                }
            }
            else
            {
                MessageBox.Show("未选择任何列的数据导出");
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
                if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
                {
                    PathText.Text = 保存地址 = folderBrowserDialog.SelectedPath;
                }
            }
            catch (Exception)
            {
            }
        }
    }

}
