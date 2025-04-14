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
                        PathText.Text = 保存地址;
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

        // 在窗体设计器中添加一个CheckBox控件，命名为 chkExportFilteredData
        // 设置其Text属性为"仅导出筛选后的数据"


        private void 文件导出_Click(object sender, EventArgs e)
        {
            
           
            // 1. 检查是否选中列
            if (CheckList.CheckedItems.Count == 0)
            {
                MessageBox.Show("请至少选择一列数据导出");
                return;
            }
            //禁用所有Excel更新和计算
            StaticClass.ExcelApp.ScreenUpdating = false;
            StaticClass.ExcelApp.Calculation = XlCalculation.xlCalculationManual;
            StaticClass.ExcelApp.DisplayAlerts = false;
            StaticClass.ExcelApp.EnableEvents = false;

            // 2. 获取选中列索引（Excel从1开始）
            List<int> selectedColumns = CheckList.Items.Cast<string>()
                .Select((item, index) => new { item, index })
                .Where(x => CheckList.GetItemChecked(x.index))
                .Select(x => x.index + 1)
                .ToList();

            Workbook sourceWorkbook = null;
            Worksheet sourceSheet = null;
            Workbook newWorkbook = null;
            Worksheet newWorksheet = null;
            Range visibleRange = null;

            try
            {
                // 3. 获取当前Excel对象
                sourceWorkbook = StaticClass.ExcelApp.ActiveWorkbook;
                sourceSheet = sourceWorkbook.ActiveSheet;
                int rows= sourceSheet.UsedRange.Rows.Count;
                // 4. 检查筛选模式
                bool exportFiltered = checkBox2.Checked;
                bool falg = sourceSheet.AutoFilter == null || !sourceSheet.AutoFilter.FilterMode;



                // 5. 创建新工作簿
                newWorkbook = StaticClass.ExcelApp.Workbooks.Add();
                newWorksheet = newWorkbook.Sheets[1];

                // 6. 复制表头（保留格式）
                //for (int i = 0; i < selectedColumns.Count; i++)
                //{
                //    Range headerCell = sourceSheet.Cells[1, selectedColumns[i]];
                //    headerCell.Copy(newWorksheet.Cells[1, i + 1]);
                //};


                //// 7. 高性能数据导出
                //if (exportFiltered && !falg)

                //{
                //    // 获取整个数据区域（包括标题）
                //    Range entireRange = sourceSheet.UsedRange;

                //    // 获取可见单元格（从第二行开始，跳过标题）
                //  visibleRange = entireRange.Offset[1, 0].Resize[entireRange.Rows.Count - 1, entireRange.Columns.Count]
                //                        .SpecialCells(XlCellType.xlCellTypeVisible);
                //    // 获取可见行号并排序
                //    List<int> visibleRows = visibleRange.Areas.Cast<Range>()
                //        .SelectMany(area => area.Rows.Cast<Range>().Select(r => r.Row))
                //        .Where(row => row > 1) // 排除标题行
                //        .OrderBy(row => row)
                //        .ToList();

                //    // 按列处理
                foreach (int colIndex in selectedColumns)
                {
                    int targetCol = selectedColumns.IndexOf(colIndex) + 1;
                    Range rng= sourceSheet.Range[sourceSheet.Cells[1, targetCol], sourceSheet.Cells[rows, targetCol]];
                    Range rng2= newWorksheet.Range[newWorksheet.Cells[1, targetCol], newWorksheet.Cells[rows, targetCol]];
                  //复制表头
                    //sourceSheet.Cells[1, colIndex].Copy(newWorksheet.Cells[1, targetCol]);
                    rng.Copy(rng2);
                    // 准备数据数组
                    //object[,] data = new object[visibleRows.Count, 1];
                    //for (int i = 0; i < visibleRows.Count; i++)
                    //{
                    //    data[i, 0] = sourceSheet.Cells[visibleRows[i], colIndex].Value2;
                    //}

                    //// 批量写入
                    //newWorksheet.Cells[2, targetCol].Resize[visibleRows.Count, 1].Value2 = data;
                    // 7.2 仅复制选中列

                }
                //}
                //else
                //{
                //    // 7.3 非筛选模式直接复制整列
                //    foreach (int colIndex in selectedColumns)
                //    {
                //        Range sourceCol = sourceSheet.Columns[colIndex];
                //        int targetCol = selectedColumns.IndexOf(colIndex) + 1;
                //        sourceCol.Copy(newWorksheet.Columns[targetCol]);
                //    }
                //}

                // 8. 设置文件名
                string workbookName = string.IsNullOrWhiteSpace(WBnameText.Text)
                    ? $"导出数据_{DateTime.Now:yyyyMMddHHmmss}"
                    : WBnameText.Text.Trim();

                if (!string.IsNullOrWhiteSpace(WSnameText.Text))
                {
                    newWorksheet.Name = WSnameText.Text.Trim();
                }

                // 9. 保存文件
                string savePath = Path.Combine(保存地址, $"{workbookName}.xlsx");
                newWorkbook.SaveAs(savePath);

                MessageBox.Show($"导出成功！\n文件已保存到：{savePath}");
            }
            catch (COMException comEx) when (comEx.Message.Contains("0x800A03EC"))
            {
                MessageBox.Show("没有可见数据可导出，请检查筛选条件");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导出失败：{ex.Message}\n建议关闭其他Excel进程后重试");
            }
            finally
            {
               // 恢复Excel设置
                    StaticClass.ExcelApp.ScreenUpdating = true;
                    StaticClass.ExcelApp.Calculation = XlCalculation.xlCalculationAutomatic;
                    StaticClass.ExcelApp.DisplayAlerts = true;
                    StaticClass.ExcelApp.EnableEvents = true;

                    // ...释放资源...
              
                // 10. 释放所有COM对象（严格按顺序）
                if (visibleRange != null) Marshal.ReleaseComObject(visibleRange);
                if (newWorksheet != null) Marshal.ReleaseComObject(newWorksheet);
                if (newWorkbook != null)
                {
                    newWorkbook.Close(false);
                    Marshal.ReleaseComObject(newWorkbook);
                }
                if (sourceSheet != null) Marshal.ReleaseComObject(sourceSheet);
                if (sourceWorkbook != null) Marshal.ReleaseComObject(sourceWorkbook);

                GC.Collect();
                GC.WaitForPendingFinalizers();
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
