using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using Button = System.Windows.Forms.Button;
using System.Diagnostics;
using Rectangle = System.Drawing.Rectangle;
using System.Collections.Concurrent;
using System.Collections;

namespace 插件.MyForm
{
    public partial class 数据对比 : Form
    {
        public 数据对比()
        {
            InitializeComponent();
            InitializeColorComboBox();



        }
        // 在类成员变量区添加
        private Color? selectColor = null; // 保存自定义颜色

        private List<Color> colors = new List<Color>();

        private void InitializeColorComboBox()
        {
            // 绑定绘制事件

            colorComboBox.SelectedIndexChanged += ColorComboBox_SelectedIndexChanged;
            colorComboBox.DrawItem += ColorComboBox_DrawItem;
            // 准备颜色数据
            colors = new List<Color>
            {
                Color.Red,
                Color.Green,
                Color.Blue,
                Color.Yellow,
                Color.Orange,
                Color.Purple,
                Color.Orchid,
                Color.Pink,
                Color.PaleGreen,
                Color.Magenta
            };

            // 绑定数据源
            colorComboBox.DataSource = colors;
            colorComboBox.DisplayMember = "Name";
        }

        private void ColorComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            int index = colorComboBox.SelectedIndex;
            if (index == colorComboBox.Items.Count - 1)
            {
                ColorDialog colorDialog = new ColorDialog();
                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    selectColor = colorDialog.Color;
                }
                else
                {
                    selectColor = null;
                }
            }
            else
            {
                selectColor = colors[index];
            }
        }
        private void ColorComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            Console.WriteLine(e.Index.ToString());
            if (e.Index < 0) return;

            var combo = sender as ComboBox;
            var color = (Color)combo.Items[e.Index];

            e.DrawBackground();

            // 绘制颜色方块
            Rectangle colorRect = new Rectangle(
                e.Bounds.X + 1,
                e.Bounds.Y + 1,
                combo.Width - 25,
                e.Bounds.Height - 4
            );
            if (e.Index == colors.Count - 1)
            {
                using (var brush = new SolidBrush(Color.White))
                {
                    e.Graphics.FillRectangle(brush, colorRect);
                }
                e.Graphics.DrawRectangle(Pens.Black, colorRect);
                e.Graphics.DrawString("更多颜色", new System.Drawing.Font("宋体", 8), Brushes.Black,
                    colorRect.X + 2, colorRect.Y + 2);
            }
            else
            {
                using (var brush = new SolidBrush(color))
                {
                    e.Graphics.FillRectangle(brush, colorRect);
                }
                e.Graphics.DrawRectangle(Pens.Black, colorRect);
            }
            e.DrawFocusRectangle();
        }
        // 使用更可靠的窗口激活API
        [DllImport("user32.dll")]
        private static extern void SwitchToThisWindow(IntPtr hWnd, bool fAltTab);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);
        private Excel.Application excelapp;
        private Range 相同Rng = null;
        private Range 不同Rng = null;

        private void 相同项_Click(object sender, EventArgs e)
        {
            try
            {
                if (标记单元格 != null)
                {
                    Marshal.ReleaseComObject(标记单元格);
                }

                清除标识.Enabled = true;
                Range combinedRange = null;
                Range combinedRange2 = null;
                excelapp.ScreenUpdating = false;
                excelapp.Calculation = XlCalculation.xlCalculationManual;
                if (相同项文本.Count < 1) return;
                foreach (Range rng in 区域一单元格)
                {
                    string value = rng.Value2?.ToString();
                    if (相同项文本.Contains(value))
                    {
                        if (combinedRange == null)
                        {
                            combinedRange = rng;
                        }
                        else
                        {
                            combinedRange = excelapp.Union(combinedRange, rng);
                        }
                    }
                }
                标记单元格 = combinedRange;
                foreach (Range rng in 区域二单元格)
                {
                    string value = rng.Value2?.ToString();
                    if (相同项文本.Contains(value))
                    {
                        if (combinedRange2 == null)
                        {
                            combinedRange2 = rng;
                        }
                        else
                        {
                            combinedRange2 = excelapp.Union(combinedRange2, rng);
                           
                        }

                    }
                }
    
                标记单元格2 = combinedRange2;
                combinedRange.Interior.Color = selectColor;
                combinedRange2.Interior.Color = selectColor;
                excelapp.ScreenUpdating = true;
                excelapp.Calculation = XlCalculation.xlCalculationAutomatic;
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void 不同项_Click(object sender, EventArgs e)
        {
            try
            {
                if (标记单元格 != null)
                {
                    Marshal.ReleaseComObject(标记单元格);
                }

                清除标识.Enabled = true;
                Range combinedRange = null;
               
                excelapp.ScreenUpdating = false;
                excelapp.Calculation = Excel.XlCalculation.xlCalculationManual;
                if (区域一独有.Count < 1 && 区域二独有.Count < 1) return;
                if (!(区域一独有.Count < 1))
                {
                    foreach (Range rng in 区域一单元格)
                    {
                        string value = rng.Value2?.ToString();
                        if (!相同项文本.Contains(value))
                        {
                            if (combinedRange == null)
                            {
                                combinedRange = rng;
                            }
                            else
                            {
                                combinedRange = excelapp.Union(combinedRange, rng);
                            }
                        }
                    }

                }
                标记单元格=combinedRange;
                Range combinedRange2 = null;
                if (!(区域二独有.Count < 1))
                {
                    foreach (Range rng in 区域二单元格)
                    {
                        string value = rng.Value2?.ToString();
                        if (!相同项文本.Contains(value))
                        {
                            if (combinedRange2 == null)
                            {
                                combinedRange2 = rng;
                            }
                            else
                            {
                                combinedRange2 = excelapp.Union(combinedRange2, rng);
                            }

                        }
                    }
                }
                标记单元格2 = combinedRange2;
                combinedRange.Interior.Color = selectColor;
                combinedRange2.Interior.Color = selectColor;
                excelapp.ScreenUpdating = true;
                excelapp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            }
            catch (Exception)
            {

                throw;
            }
        }
        private Range 标记单元格 = null;
        private Range 标记单元格2 = null;

        private void 清除标识_Click(object sender, EventArgs e)
        {
            try
            {
                  excelapp.ScreenUpdating = false;
                excelapp.Calculation = XlCalculation.xlCalculationManual;

                标记单元格.Interior.Color = XlColorIndex.xlColorIndexNone;



                // 关闭屏幕更新和自动计算
                excelapp.ScreenUpdating = true;
                excelapp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
                Button button = (Button)sender;
                button.Enabled = false;
            }
            catch (Exception)
            {

                throw;
            }

        }

        private void 导出相同项_Click(object sender, EventArgs e)
        {
            IntPtr excelHandle = new IntPtr(excelapp.Hwnd);
            SetForegroundWindow(excelHandle);
            SwitchToThisWindow(excelHandle, true);
            this.Hide();
            Range rng = (Range)excelapp.InputBox(Prompt: "请选择单元格", Title: "选择单元格", Default: "", Type: 8);
            Excel.Range targetRange = rng.Resize[相同项文本.Count, 1];
            // 将一维数组转换为二维数组
            object[,] dataArray = new object[相同项文本.Count, 1];
            for (int i = 0; i < 相同项文本.Count; i++)
            {
                dataArray[i, 0] = 相同项文本[i];
            }

            // 将二维数组赋值给目标范围
            targetRange.Value2 = dataArray;
            this.Show();
        }

        private void 导出不同项_Click(object sender, EventArgs e)
        {

            IntPtr excelHandle = new IntPtr(excelapp.Hwnd);
            SetForegroundWindow(excelHandle);
            SwitchToThisWindow(excelHandle, true);
            this.Hide();
            Range rng = (Range)excelapp.InputBox(Prompt: "请选择单元格", Title: "选择单元格", Default: "", Type: 8);
            区域一独有.AddRange(区域二独有);
            Excel.Range targetRange = rng.Resize[区域二独有.Count, 1];
            // 将一维数组转换为二维数组
            object[,] dataArray = new object[区域二独有.Count, 1];
            for (int i = 0; i < 区域二独有.Count; i++)
            {
                dataArray[i, 0] = 区域二独有[i];
            }

            targetRange.Value2 = dataArray;
            this.Show();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            相同项.Enabled = false;
            不同项.Enabled = false;
            清除标识.Enabled = false;
            导出不同项.Enabled = false;
            导出相同项.Enabled = false;
            KeyPreview = true;
            KeyDown += Form2_KeyDown;
            excelapp = StaticClass.ExcelApp;
        }

        private void Form2_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.Escape:
                    break;
                case Keys.S:
                    相同项.PerformClick();
                    break;
                case Keys.D:
                    e.SuppressKeyPress = false;
                    不同项.PerformClick();
                    break;
                case Keys.C:
                    e.SuppressKeyPress = false;
                    清除标识.PerformClick();
                    break;
                case Keys.E:
                    e.SuppressKeyPress = false;
                    导出相同项.PerformClick();
                    break;
                case Keys.F:
                    e.SuppressKeyPress = false;
                    导出不同项.PerformClick();
                    break;
                default:
                    break;
            }
            this.Focus();
        }

        private Range 区域一单元格 = null;
        private Range 区域二单元格 = null;

        private List<string> 区域一独有 = new List<string>();
        private List<string> 区域二独有 = new List<string>();
        private List<string> 相同项文本 = new List<string>();
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                IntPtr excelHandle = new IntPtr(excelapp.Hwnd);
                SetForegroundWindow(excelHandle);
                SwitchToThisWindow(excelHandle, true);
                this.Hide();
                // 使用 InputBox 方法提示用户选择单元格
                object result = excelapp.InputBox(Prompt: "请选择单元格", Title: "选择单元格", Default: "", Type: 8); // Type 8 表示返回一个 Range 对象                                                                                                      
                if (result != null)  // 检查用户是否取消了选择
                {
                    if (result is Range selectedRange)
                    {
                        Worksheet selectedWorksheet = selectedRange.Worksheet;
                        worksheet1 = selectedWorksheet;
                        // 获取工作簿
                        Workbook selectedWorkbook = selectedWorksheet.Parent;
                        int selectIndex = selectedRange.Rows.Count;
                        int RowsCount = selectedWorksheet.Rows.Count;
                        int rowIndex = selectIndex;
                        if (selectIndex == RowsCount)
                        {
                            Range r = selectedWorksheet.Cells[selectedWorksheet.Rows.Count, selectedRange.Column];
                            rowIndex = r.End[XlDirection.xlUp].Row;
                        }
                        int columnCount = selectedRange.Columns.Count;
                        Range rng = selectedWorksheet.Range[selectedRange.Cells[1, 1], selectedRange.Cells[rowIndex, columnCount]];
                        区域一单元格 = rng;
                        区域1Box.Text = BuildSmartAddress(excelapp, selectedWorkbook, selectedWorksheet, selectedRange);
                        
                    }
                }
                this.Show();
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            try
            {
                IntPtr excelHandle = new IntPtr(excelapp.Hwnd);
                SetForegroundWindow(excelHandle);
                SwitchToThisWindow(excelHandle, true);
                this.Hide();
                // 使用 InputBox 方法提示用户选择单元格
                object result = excelapp.InputBox(Prompt: "请选择单元格", Title: "选择单元格", Default: "", Type: 8); // Type 8 表示返回一个 Range 对象
                                                                                                           // 检查用户是否取消了选择
                if (result != null)
                {
                    // 将结果转换为 Range 对象

                    if (result is Range selectedRange)
                    {
                        Worksheet selectedWorksheet = selectedRange.Worksheet;
                        worksheet2 = selectedWorksheet;
                        // 获取工作簿
                        Workbook selectedWorkbook = selectedWorksheet.Parent;
                        int selectIndex = selectedRange.Rows.Count;
                        int RowsCount = selectedWorksheet.Rows.Count;
                        int rowIndex = selectIndex;
                        if (selectIndex == RowsCount)
                        {
                            Range r = selectedWorksheet.Cells[selectedWorksheet.Rows.Count, selectedRange.Column];
                            rowIndex = r.End[XlDirection.xlUp].Row;
                        }

                        int columnCount = selectedRange.Columns.Count;
                        Range rng = selectedWorksheet.Range[selectedRange.Cells[1, 1], selectedRange.Cells[rowIndex, columnCount]];
                        区域二单元格 = rng;
                        区域2Box.Text = BuildSmartAddress(excelapp, selectedWorkbook, selectedWorksheet, selectedRange);

                    }
                }
                this.Show();
            }
            catch (Exception)
            {

                throw;
            }
        }
        private Worksheet worksheet1 = null;
        private Worksheet worksheet2 = null;
        /// <summary>
        /// 智能生成地址表示（自动省略相同工作簿/工作表信息）
        /// </summary>
        /// 
        private string BuildSmartAddress(Excel.Application excelApp,
                                      Workbook workbook,
                                     Worksheet worksheet,
                                       Range range)
        {
            var address = range.Address[Excel.XlReferenceStyle.xlA1]
                           .Replace("$", "");

            // 当前工作簿
            if (workbook == excelApp.ActiveWorkbook)
            {
                // 当前工作表
                if (worksheet == excelApp.ActiveSheet)
                    return address;

                return $"{worksheet.Name}!{address}";
            }

            return $"[{workbook.Name}]{worksheet.Name}!{address}";
        }
        //保存相同和不同的单元格地址
        // 使用线程安全的集合
        ConcurrentBag<string> 数据1 = new ConcurrentBag<string>();

        ConcurrentBag<string> 数据2 = new ConcurrentBag<string>();

        private void 对比数据_Click(object sender, EventArgs e)
        {
            try
            {
                if (数据1.Count>0)
                {
                    string item;
                    while (数据1.TryTake(out item))
                    {
                     
                    }
                }
                if (数据2.Count > 0) {
                    string item;
                    while (数据2.TryTake(out item))
                    {

                    }
                }
                if (区域一独有.Count > 0)
                {
                    区域一独有.Clear();
                }
                if (区域二独有.Count > 0)
                {
                    区域二独有.Clear();
                }
                if (相同项文本.Count > 0)
                {
                    相同项文本.Clear();
                }

                相同项Text.Text = "";
                区域一Text.Text = "";
                区域二Text.Text = "";
                if (string.IsNullOrEmpty(区域1Box.Text) || string.IsNullOrEmpty(区域2Box.Text))
                {
                    MessageBox.Show("请先选择两个对比区域");
                    return;
                }
                // 使用快速数组访问
                object[,] data1 = (object[,])区域一单元格.Value;
                object[,] data2 = (object[,])区域二单元格.Value;
                if (data1 == null || data2 == null) return;
                CountdownEvent countdownEvent = new CountdownEvent(2);

                // 启动任务处理 data1
                ThreadPool.QueueUserWorkItem(_ =>
                {
                    try
                    {
                        获取单元格数据(data1, 数据1);
                    }
                    finally
                    {
                        // 任务完成，发出信号
                        countdownEvent.Signal();
                    }
                });
                // 启动任务处理 data2
                ThreadPool.QueueUserWorkItem(_ =>
                {
                    try
                    {
                        获取单元格数据(data2, 数据2);
                    }
                    finally
                    {
                        // 任务完成，发出信号
                        countdownEvent.Signal();
                    }
                });
                countdownEvent.Wait();//把区域的所有数据提取完成
                相同项文本 = 数据1.Intersect(数据2).ToList();//求出相同项
           if (相同项文本 != null)
                {

                    区域一独有 = 数据1.Where(str => !数据2.Contains(str)).ToList();
                    区域二独有 = 数据2.Where(str => !数据1.Contains(str)).ToList();
                }
                相同项.Enabled = 导出相同项.Enabled = 相同项文本.Count > 1;
                不同项.Enabled = 导出不同项.Enabled = 区域二独有.Count > 1 || 区域一独有.Count > 1;
                区域一Text.Text = string.Join(Environment.NewLine, 区域一独有.ToArray());
                区域二Text.Text = string.Join(Environment.NewLine, 区域二独有.ToArray());
                相同项Text.Text = string.Join(Environment.NewLine, 相同项文本.ToArray());
            }
            catch (Exception ex)
            {
                MessageBox.Show($"处理失败: {ex.Message}");
            }
        }

        private void 获取单元格数据(object[,] data, ConcurrentBag<string> 数据)
        {
            try
            {
                for (int i = 1; i <= data.GetLength(1); i++)
                {
                    for (int j = 1; j <= data.GetLength(0); j++)
                    {
                        object o = data[j, i];
                        string str = o.ToString();
                        if (!数据.Contains(str))
                        { 
                            数据.Add(str);
                        }                
                    }
                }
            }
            catch (Exception ex) { }
        }
        private void ReleaseComObjects(List<Excel.Range> ranges)
        {
            foreach (var range in ranges)
            {
                if (range != null && Marshal.IsComObject(range))
                {
                    Marshal.ReleaseComObject(range);
                }
            }
            System.GC.Collect(); // 强制回收释放的COM对象
        }

        private void 退出_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void 数据对比_FormClosed(object sender, FormClosedEventArgs e)
        {

            Globals.ThisAddIn.对比form = null;
        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }
    }
}