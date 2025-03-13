using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
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
            try
            {


                worksheet = (Worksheet)excelapp.ActiveSheet;
                后缀com.Items.AddRange(new object[] { ".xlsx", ".xlsm", ".txt", ".xls" });
                后缀com.SelectedIndex = 0;
                string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                桌面路径.Text = desktopPath;
                GETguanjianzi();
                表头行数 = 1;
                关键名com.Items.AddRange(new object[] { "1月份", "2月份", "3月份", "4月份", "5月份", "6月份", "7月份", "8月份", "9月份", "10月份", "11月份", "12月份" });
            }
            catch (Exception)
            {

                throw;
            }
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
            try
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
            catch (Exception)
            {

                throw;
            }
        }

        private void 关键列com_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                关键字com.Items.Clear();
                int rows = worksheet.UsedRange.Rows.Count;
                HashSet<string> list = new HashSet<string>(rows);
                string str = 关键列com.SelectedItem.ToString();
                SelectCol = str.Split(':')[0];
                Range range = worksheet.Range[$"{SelectCol}2:{SelectCol}{rows}"];
                if (range != null)
                {
                    关键字com.Items.Add("All");

                    foreach (Range r in range)
                    {
                        if (r.Value2 != null)
                        {
                            list.Add(r.Value2.ToString());
                        }
                    }
                    关键字com.Items.AddRange(list.ToArray());
                    关键字com.SelectedIndex = 0;
                }
                vlaue = list;

            }
            catch (Exception)
            {
            }
        }
 
        
        private void 拆分()
        {
            Workbook summaryWorkbook = null;
            Worksheet summarySheet = null;

            try
            {
                excelapp.Visible = false;
                // 禁用屏幕刷新和事件触发
                excelapp.ScreenUpdating = false;
                excelapp.EnableEvents = false;

                // 创建汇总工作簿
                if (关键名com.SelectedIndex == 0)
                {
                    summaryWorkbook = excelapp.Workbooks.Add();
                    summarySheet = (Worksheet)summaryWorkbook.Sheets[1];
                    summarySheet.Name = "超链接汇总";
                }
                else
                {
                    summaryWorkbook = excelapp.ActiveWorkbook;
                    summarySheet = summaryWorkbook.Sheets.Add();
                    summarySheet.Name = "超链接汇总";
                }
                    int summaryRow = 1;

                // 确保目录存在
                string basePath = Path.Combine(桌面路径.Text, 关键名com.Text, textBox1.Text, textBox2.Text);
                Directory.CreateDirectory(basePath);

                // 获取关键字集合
                HashSet<string> names = 关键字com.SelectedIndex> 0
                    ? new HashSet<string> { 关键字com.SelectedItem.ToString()}
                    : vlaue;

                // 遍历每个关键字
                foreach (string keyword in names)
                {
                    Workbook newWorkbook = null;
                    Worksheet newSheet = null;
                    try
                    {
                        // 创建一个新的工作簿
                        newWorkbook = excelapp.Workbooks.Add();
                        newSheet = (Worksheet)newWorkbook.Sheets[1];
                        newSheet.Name = keyword;

                        // 复制表头
                        Range headerRange = worksheet.Rows[表头行数];
                        headerRange.Copy(newSheet.Rows[表头行数]);

                        // 遍历数据行并复制匹配的行
                        int newRow = 表头行数 + 1;
                        int rows = worksheet.UsedRange.Rows.Count;
                        for (int i = 2; i <= rows; i++)
                        {
                            Range cell = worksheet.Cells[i, SelectCol];
                            try
                            {
                                if (cell.Value2 != null && cell.Value2.ToString() == keyword)
                                {
                                    Range rowRange = worksheet.Rows[i];
                                    CopyRowWithOptions(rowRange, newSheet.Rows[newRow]);
                                    newRow++;
                                }
                            }
                            finally
                            {
                                if (cell != null) Marshal.ReleaseComObject(cell);
                            }
                        }
                        // 保存单独的工作簿
                        string filePath = Path.Combine(basePath, $"{keyword}{后缀com.Text}");
                        newWorkbook.SaveAs(filePath);
                        newWorkbook.Close(false); // 不保存更改提示

                        // 在汇总工作表中添加超链接
                        summarySheet.Hyperlinks.Add(
                            Anchor: summarySheet.Cells[summaryRow, 1],
                            Address: filePath,
                            TextToDisplay: keyword
                        );
                        summaryRow++;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"处理关键字 {keyword} 时发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Console.WriteLine($"Error processing keyword {keyword}: {ex.Message}");
                    }
                    finally
                    {
                        // 释放当前工作簿和工作表对象
                        if (newSheet != null) Marshal.FinalReleaseComObject(newSheet);
                        if (newWorkbook != null) Marshal.FinalReleaseComObject(newWorkbook);

                        // 强制垃圾回收
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                }

                // 保存汇总工作簿
                if (关键名com.SelectedIndex==0)
                {

                    string summaryFilePath = Path.Combine(Path.GetDirectoryName(basePath), $"汇总{后缀com.Text}");
                    summaryWorkbook.SaveAs(summaryFilePath);
                    summaryWorkbook.Close(false); // 不保存更改提示
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Console.WriteLine($"Error: {ex.Message}");
            }
            finally
            {
                // 确保释放所有 COM 对象
                if (summarySheet != null) Marshal.FinalReleaseComObject(summarySheet);
                if (summaryWorkbook != null) Marshal.FinalReleaseComObject(summaryWorkbook);

                // 强制垃圾回收
                GC.Collect();
                GC.WaitForPendingFinalizers();

              

                // 恢复屏幕刷新和事件触发
                excelapp.ScreenUpdating = true;
                excelapp.EnableEvents = true;
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

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }
    }
}