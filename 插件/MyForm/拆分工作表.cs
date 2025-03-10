using Microsoft.Office.Interop.Excel;
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

namespace 插件.MyForm
{
    public partial class 拆分工作表 : Form
    {
        public 拆分工作表()
        {
            InitializeComponent();
        }
        private readonly Microsoft.Office.Interop.Excel.Application excelapp = StaticClass.ExcelApp;
        private int 表头行数 = 0;
        private Worksheet worksheet;
        private HashSet<string> vlaue;
        private string SelectCol;

        private void 拆分工作表_Load(object sender, EventArgs e)
        {
            worksheet = (Worksheet)excelapp.ActiveSheet;
            后缀com.Items.AddRange(new object[] { ".xlsx", ".xlsm", ".txt", ".xls" });
            后缀com.SelectedIndex = 0;
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            桌面路径.Text = desktopPath;
            GETguanjianzi();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            拆分();
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            表头行数 = (int)numericUpDown1.Value;
        }

        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox4.Checked)
            {
                checkBox1.Checked = true;
                checkBox2.Checked = true;
                checkBox3.Checked = true;
            }
            else
            {
                checkBox1.Checked = false;
                checkBox2.Checked = false;
                checkBox3.Checked = false;
            }
        }

        private void GETguanjianzi()
        {
            // 找到第一行第一个有内容的列
            int firstNonEmptyCol = 1;
            Range firstRow = worksheet.Rows[1];
            // 获取第一行最后一个有内容的列

            foreach (Range cell in firstRow.Cells)
            {
                if (cell.Value2 != null)
                {
                    firstNonEmptyCol = cell.Column;
                    break;
                }
            }
            Range lastCell = worksheet.Cells[1, worksheet.Columns.Count];
            int lastColumn = lastCell.End[XlDirection.xlToLeft].Column;
            Range rng = worksheet.Range[worksheet.Cells[1, firstNonEmptyCol], worksheet.Cells[1, lastColumn]];
            if (rng != null)
            {
                foreach (Range r in rng)
                {
                    string str = r.Address[false, false].Replace("1", "");

                    关键列com.Items.Add($"{str}:{r.Value2}");
                }
                关键列com.SelectedIndex = 0;
            }
        }

        private void 关键列com_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                关键字con.Items.Clear();
                int rows = worksheet.UsedRange.Rows.Count;
                HashSet<string> list = new HashSet<string>(rows);
                string str = 关键列com.SelectedItem.ToString();
                SelectCol = str.Split(':')[0];
                Range range = worksheet.Range[$"{SelectCol}2:{SelectCol}{rows}"];
                if (range != null)
                {
                    foreach (Range r in range)
                    {
                        if (r.Value2 != null)
                        {
                            list.Add(r.Value2.ToString());
                        }
                    }
                    关键字con.Items.AddRange(list.ToArray());
                    关键字con.SelectedIndex = 0;
                }
                vlaue = list;
               
            }
            catch (Exception)
            {
            }
        }

        private void 拆分()
        {
            try
            {// 创建汇总工作簿
               Workbook summaryWorkbook = excelapp.Workbooks.Add();
                Worksheet summarySheet = (Worksheet)summaryWorkbook.Sheets[1];
                summarySheet.Name = "超链接汇总";
                int summaryRow = 1;

                // 确保目录存在
                string basePath = Path.Combine(桌面路径.Text, textBox1.Text, textBox2.Text);
                Directory.CreateDirectory(basePath);

                // 遍历每个关键字
                foreach (string keyword in vlaue)
                {
                    // 创建一个新的工作簿
                 Workbook newWorkbook = excelapp.Workbooks.Add();
                  Worksheet newSheet = (Worksheet)newWorkbook.Sheets[1];
                    newSheet.Name = keyword;

                    // 复制表头
                    Range headerRange = worksheet.Rows[表头行数];
                    headerRange.Copy(newSheet.Rows[表头行数]);

                    int newRow = 表头行数 + 1; // 用于记录新工作表中的行号
                    int rows = worksheet.UsedRange.Rows.Count;
                    for (int i = 2; i <= rows; i++)
                    {
                       Range cell = worksheet.Cells[i, SelectCol];
                        if (cell.Value2 != null && cell.Value2.ToString() == keyword)
                        {
                            Range rowRange = worksheet.Rows[i];
                            CopyRowWithOptions(rowRange, newSheet.Rows[newRow]);
                            newRow++;
                        }
                        Marshal.ReleaseComObject(cell);
                    }

                    // 保存单独的工作簿
                    string filePath = Path.Combine(basePath, $"{keyword}{后缀com.Text}");
                    newWorkbook.SaveAs(filePath);
                    newWorkbook.Close();

                    // 释放新工作簿和工作表对象
                    Marshal.ReleaseComObject(newSheet);
                    Marshal.ReleaseComObject(newWorkbook);

                    // 在汇总工作表中添加超链接
                    summarySheet.Hyperlinks.Add(Anchor: summarySheet.Cells[summaryRow, 1], Address: filePath, TextToDisplay: keyword);
                    summaryRow++;
                }

                // 保存汇总工作簿
                string summaryFilePath = Path.Combine(basePath, $"汇总{后缀com.Text}");
                summaryWorkbook.SaveAs(summaryFilePath);
                summaryWorkbook.Close();

                // 释放汇总工作簿和工作表对象
                Marshal.ReleaseComObject(summarySheet);
                Marshal.ReleaseComObject(summaryWorkbook);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // 可以添加更详细的日志记录
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
        private void CopyRowWithOptions(Range sourceRow, Range targetRow)
        {
            if (checkBox4.Checked)
            {
                sourceRow.Copy(targetRow);
            }
            else
            {
                if (checkBox1.Checked)
                {
                    sourceRow.Copy();
                    targetRow.PasteSpecial(XlPasteType.xlPasteFormats);
                }
                if (checkBox2.Checked)
                {
                    sourceRow.Copy();
                    targetRow.PasteSpecial(XlPasteType.xlPasteComments);
                }
                if (checkBox3.Checked)
                {
                    targetRow.Formula = sourceRow.Formula;
                }
                if (!checkBox1.Checked && !checkBox3.Checked && !checkBox2.Checked)
                {
                    targetRow.Value = sourceRow.Value;
                }
            }
        }

        private void 浏览_Click(object sender, EventArgs e)
        {
            try
            { 
               FolderBrowserDialog dialog = new FolderBrowserDialog();
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    桌面路径.Text = dialog.SelectedPath;
                }
            }
            catch (Exception)
            {
       
            }


        }

        private void 拆分工作表_FormClosed(object sender, FormClosedEventArgs e)
        {
            AMG.拆分form = null;
        }
    }
}