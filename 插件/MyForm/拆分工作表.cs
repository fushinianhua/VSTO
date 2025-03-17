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
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Excel.Application;

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

        /// <summary>
        /// 
        /// </summary>
        private void 拆分()
        {
           Application newExcelApp = null;
         
            Workbook currentWorkbook = excelapp.ActiveWorkbook;
            Worksheet summarySheet = null;
            Workbook summaryWorkbook = null;
            int summaryRow = 1;
            try
            {
                // 初始化新Excel实例
                newExcelApp = new Application();
                newExcelApp.Visible = false;
                newExcelApp.ScreenUpdating = false;
                newExcelApp.DisplayAlerts = false;

                // 读取原始数据
                Range usedRange = worksheet.UsedRange;
                object[,] allData = (object[,])usedRange.Value2;
                int totalRows = usedRange.Rows.Count;
                int totalCols = usedRange.Columns.Count;

                // 释放Range对象
                Marshal.ReleaseComObject(usedRange);
                usedRange = null;

                // 处理模式判断
                bool isMultiMode = 关键字com.SelectedIndex == 0;
                string targetKeyword = isMultiMode ? null : 关键字com.SelectedItem.ToString();

                // 构建基础路径
                string baseFolder = Path.Combine(桌面路径.Text, 关键名com.Text, textBox1.Text, textBox2.Text);
                Directory.CreateDirectory(baseFolder);

                // 初始化汇总表
                if (isMultiMode)
                {
                    summaryWorkbook = newExcelApp.Workbooks.Add();
                    summarySheet = (Worksheet)summaryWorkbook.Sheets[1];
                    summarySheet.Name = "汇总";
                }
                else
                {
                    List<string> names = new List<string>();
                    foreach (Worksheet ws in currentWorkbook.Sheets)
                    {
                        names.Add(ws.Name);
                    }
                    if (names.Contains(targetKeyword))
                    {
                        summarySheet = (Worksheet)currentWorkbook.Worksheets[targetKeyword];
                        Range rng = (Range)summarySheet.Cells[1,summarySheet.Rows.Count];
                        summaryRow = rng.End[XlDirection.xlDown].Row+1;
                    }
                    else
                    {
                        summarySheet = (Worksheet)currentWorkbook.Sheets.Add(
                         After: currentWorkbook.Sheets[currentWorkbook.Sheets.Count]);
                        summarySheet.Name = targetKeyword;
                    }
                }
              
                // 数据收集
                var dataDict = new Dictionary<string, List<object[,]>>();
                for (int row = 2; row <= totalRows; row++)
                {
                    var cellValue = allData[row, 关键列com.SelectedIndex+1];
                    if (cellValue == null) continue;

                    string key = cellValue.ToString();
                    if (!isMultiMode && key != targetKeyword) continue;

                    if (!dataDict.ContainsKey(key))
                        dataDict[key] = new List<object[,]>();

                    object[,] rowData = new object[1, totalCols];
                    Array.Copy(allData, (row - 1) * totalCols, rowData, 0, totalCols);
                    dataDict[key].Add(rowData);
                }
                // 文件生成逻辑          
                foreach (var kv in dataDict)
                {
                    string keyword = kv.Key;
                    string safeKeyword = CleanFileName(keyword);
                    string filePath = Path.Combine(baseFolder, $"{safeKeyword}{后缀com.Text}");

                    Workbook newWorkbook = null;
                    Worksheet newWorksheet = null;
                    Range dataRange = null;

                    try
                    {
                        // 创建新工作簿
                        newWorkbook = newExcelApp.Workbooks.Add();
                        newWorksheet = (Worksheet)newWorkbook.Sheets[1];
                        newWorksheet.Name = safeKeyword;

                        // 准备数据
                        int dataCount = kv.Value.Count;
                        object[,] outputData = new object[dataCount + 1, totalCols];
                        Array.Copy(allData, outputData, totalCols); // 复制表头

                        for (int i = 0; i < dataCount; i++)
                        {
                            Array.Copy(kv.Value[i], 0, outputData, (i + 1) * totalCols, totalCols);
                        }

                        // 批量写入
                        dataRange = newWorksheet.Range[
                            newWorksheet.Cells[1, 1],
                            newWorksheet.Cells[dataCount + 1, totalCols]];
                        dataRange.Value2 = outputData;

                        // 保存文件
                        newWorkbook.SaveAs(filePath);
                    }
                    finally
                    {
                        // 释放资源
                        if (dataRange != null) Marshal.ReleaseComObject(dataRange);
                        if (newWorksheet != null) Marshal.ReleaseComObject(newWorksheet);
                        if (newWorkbook != null)
                        {
                            newWorkbook.Close(false);
                            Marshal.ReleaseComObject(newWorkbook);
                        }
                    }

                    // 添加超链接
                    summarySheet.Hyperlinks.Add(
                        summarySheet.Cells[summaryRow, 1],
                        filePath,
                        TextToDisplay: $"{keyword}"
                    );
                    summaryRow++;
                }

                // 保存汇总
                if (isMultiMode)
                {
                    string summaryPath = Path.Combine(
                        Path.GetDirectoryName(baseFolder),
                        $"数据汇总{后缀com.Text}");
                    summaryWorkbook.SaveAs(summaryPath);
                }         
            }
            catch (Exception ex)
            {
                MessageBox.Show($"操作失败：{ex.Message}");
            }
            finally
            {
                // 释放所有COM对象
                if (summarySheet != null) Marshal.ReleaseComObject(summarySheet);
                if (summaryWorkbook != null)
                {
                    summaryWorkbook.Close(false);
                    Marshal.ReleaseComObject(summaryWorkbook);
                }
                if (newExcelApp != null)
                {
                    newExcelApp.Quit();
                    Marshal.ReleaseComObject(newExcelApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // 清理非法文件名字符
        private string CleanFileName(string fileName)
        {
            return Regex.Replace(fileName, @"[\\/:*?""<>|]", "_");
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

            Globals.ThisAddIn.拆分form = null;
        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }
    }
}