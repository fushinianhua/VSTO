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
                using (var brush = new SolidBrush( Color.White))
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



        private void 相同项_Click(object sender, EventArgs e)
        {
            清除标识.Enabled = true;
            Range combinedRange = null;
            excelapp.ScreenUpdating = false;
            excelapp.Calculation = Excel.XlCalculation.xlCalculationManual;
            foreach (Range range in 相同Rng)
            {
                if (combinedRange == null)
                {
                    combinedRange = range;
                }
                else
                {
                    combinedRange = excelapp.Union(combinedRange, range);
                }
            }
            combinedRange.Interior.Color = selectColor;
            excelapp.ScreenUpdating = true;
            excelapp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
        }

        private void 不同项_Click(object sender, EventArgs e)
        {

            清除标识.Enabled = true;
            Range combinedRange = null;
            excelapp.ScreenUpdating = false;
            excelapp.Calculation = Excel.XlCalculation.xlCalculationManual;
            foreach (Range range in 不同Rng)
            {
                if (combinedRange == null)
                {
                    combinedRange = range;
                }
                else
                {
                    combinedRange = excelapp.Union(combinedRange, range);
                }
            }
            combinedRange.Interior.Color = selectColor;
            excelapp.ScreenUpdating = true;
            excelapp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
        }

        private void 清除标识_Click(object sender, EventArgs e)
        {
            Range combinedRange = null;
            // 关闭屏幕更新和自动计算
            excelapp.ScreenUpdating = false;
            excelapp.Calculation = Excel.XlCalculation.xlCalculationManual;
            foreach (Range range in 不同Rng)
            {
                if (combinedRange == null)
                {
                    combinedRange = range;
                }
                else
                {
                    combinedRange = excelapp.Union(combinedRange, range);
                }
            }
            foreach (Range range in 相同Rng)
            {
                if (combinedRange == null)
                {
                    combinedRange = range;
                }
                else
                {
                    combinedRange = excelapp.Union(combinedRange, range);
                }
            }

            combinedRange.Interior.Color = XlColorIndex.xlColorIndexNone;

            // 关闭屏幕更新和自动计算
            excelapp.ScreenUpdating = true;
            excelapp.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            Button button = (Button)sender;
            button.Enabled = false;
        }

        private void 导出相同项_Click(object sender, EventArgs e)
        {
            IntPtr excelHandle = new IntPtr(excelapp.Hwnd);
            SetForegroundWindow(excelHandle);
            SwitchToThisWindow(excelHandle, true);
            this.Hide();
            Range rng = (Range)excelapp.InputBox(Prompt: "请选择单元格", Title: "选择单元格", Default: "", Type: 8);
            rng.Resize[commonKeys.Count].Value2 = 相同Rng.ToArray();
            this.Show();
        }

        private void 导出不同项_Click(object sender, EventArgs e)
        {

            IntPtr excelHandle = new IntPtr(excelapp.Hwnd);
            SetForegroundWindow(excelHandle);
            SwitchToThisWindow(excelHandle, true);
            this.Hide();
            Range rng = (Range)excelapp.InputBox(Prompt: "请选择单元格", Title: "选择单元格", Default: "", Type: 8);
            rng.Resize[uniqueKeys1.Count+uniqueKeys2.Count].Value2=不同Rng.ToArray();
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

        private Range 区域一 = null;
        private Range 区域二 = null;


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
                    if (result is Excel.Range selectedRange)
                    {
                        // 获取选择区域所在的工作表
                        Excel.Worksheet selectedWorksheet = selectedRange.Worksheet;
                        // 获取工作簿
                        Excel.Workbook selectedWorkbook = selectedWorksheet.Parent;
                        // 构建完整的地址信息，格式为 [工作簿名称]工作表名称!单元格地址
                        区域一 = selectedRange;
                        区域1Box.Text = BuildSmartAddress(excelapp, selectedWorkbook, selectedWorksheet, selectedRange); ;
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

                    if (result is Excel.Range selectedRange)
                    {
                        区域二 = selectedRange;
                        // 获取选择区域所在的工作表
                        Worksheet selectedWorksheet = selectedRange.Worksheet;
                        // 获取工作簿
                        Workbook selectedWorkbook = selectedWorksheet.Parent;
                        // 构建完整的地址信息，格式为 [工作簿名称]工作表名称!单元格地址
                        区域二 = selectedRange;
                        区域2Box.Text = BuildSmartAddress(excelapp, selectedWorkbook, selectedWorksheet, selectedRange); ;
                    }
                }
                this.Show();
            }
            catch (Exception)
            {

                throw;
            }
        }
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
        private List<Range> 相同Rng = new List<Range>();
        private List<Range> 不同Rng = new List<Range>();
        private HashSet<string> commonKeys;
        private HashSet<string> uniqueKeys1;
        private HashSet<string> uniqueKeys2;

        private void 对比数据_Click(object sender, EventArgs e)
        {
            try
            {
                // 重置所有集合
                相同Rng.Clear();

                不同Rng.Clear();

                // 基础验证（保持不变）
                if (string.IsNullOrEmpty(区域1Box.Text) || string.IsNullOrEmpty(区域2Box.Text))
                {
                    MessageBox.Show("请先选择两个对比区域");
                    return;
                }

                // 获取数据（保持不变）
                object[,] data1 = 区域一.Value2 as object[,];
                object[,] data2 = 区域二.Value2 as object[,];
                if (data1 == null || data2 == null) return;

                // 构建值字典（包含单元格地址）
                var dict1 = BuildValueDictionary(data1, 区域一);
                var dict2 = BuildValueDictionary(data2, 区域二);

                // 计算三个数据集
                 commonKeys = new HashSet<string>(dict1.Keys.Intersect(dict2.Keys));
               uniqueKeys1 = new HashSet<string>(dict1.Keys.Except(dict2.Keys));
                 uniqueKeys2 = new HashSet<string>(dict2.Keys.Except(dict1.Keys));

                // 填充结果集合
                foreach (var key in commonKeys)
                {
                    相同Rng.AddRange(dict1[key]); // 区域一的单元格
                    相同Rng.AddRange(dict2[key]); // 区域二的单元格
                }

                foreach (var key in uniqueKeys1)
                    不同Rng.AddRange(dict1[key]);

                foreach (var key in uniqueKeys2)
                    不同Rng.AddRange(dict2[key]);

                // 更新界面
                UpdateDisplay(commonKeys, uniqueKeys1, uniqueKeys2);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        // 构建值到单元格的映射字典
        private Dictionary<string, List<Excel.Range>> BuildValueDictionary(object[,] data, Excel.Range baseRange)
        {
            var dict = new Dictionary<string, List<Excel.Range>>();

            for (int row = 1; row <= data.GetLength(0); row++)
            {
                for (int col = 1; col <= data.GetLength(1); col++)
                {
                    var value = data[row, col];
                    var key = ConvertValueKey(value);
                    var cell = baseRange.Cells[row, col];

                    if (!dict.ContainsKey(key))
                        dict[key] = new List<Excel.Range>();

                    dict[key].Add(cell);
                }
            }
            return dict;
        }// 值标准化处理
        private string ConvertValueKey(object value)
        {
            if (value == null) return "∅"; // 特殊符号表示null
            if (value is string str) return str.Trim();
            return Convert.ToString(value);
        }
        // 界面更新
        private void UpdateDisplay(
            HashSet<string> commonKeys,
            HashSet<string> uniqueKeys1,
            HashSet<string> uniqueKeys2)
        {
            // 值显示（自动去重）
            区域一Text.Text = string.Join(Environment.NewLine, uniqueKeys1);
            区域二Text.Text = string.Join(Environment.NewLine, uniqueKeys2);
            相同项Text.Text = string.Join(Environment.NewLine, commonKeys);

            // 按钮状态控制
            不同项.Enabled = 导出不同项.Enabled = uniqueKeys1.Count > 0 || uniqueKeys2.Count > 0;
            相同项.Enabled = 导出相同项.Enabled = commonKeys.Count > 0;

            // 调试信息（可选）
            Debug.WriteLine($"相同单元格数：{相同Rng.Count}");
            Debug.WriteLine($"区域一不同单元格数：{不同Rng.Count}");
          
        }

        // 辅助方法：填充值集合
        private void FillValueSet(object[,] values, HashSet<object> set)
        {
            for (int row = 1; row <= values.GetLength(0); row++)
                for (int col = 1; col <= values.GetLength(1); col++)
                    set.Add(values[row, col]);
        }

        // 辅助方法：构建值字典
        private void BuildValueDictionary(object[,] values, Excel.Range range, Dictionary<object, List<Excel.Range>> dict)
        {
            for (int row = 1; row <= values.GetLength(0); row++)
            {
                for (int col = 1; col <= values.GetLength(1); col++)
                {
                    // 统一处理值的类型和 null
                    object rawValue = values[row, col];
                    object keyValue = rawValue == null ? null : Convert.ChangeType(rawValue, typeof(string));

                    // 记录原始单元格
                    Excel.Range cell = range.Cells[row, col];

                    if (!dict.ContainsKey(keyValue))
                        dict[keyValue] = new List<Excel.Range>();

                    dict[keyValue].Add(cell); // 直接存储区域二的单元格
                }
            }
        }
        // 辅助方法：查找匹配项
        // 修改后的 FindMatches 方法
        private void FindMatches(
        object[,] values,
        Excel.Range range,
        Dictionary<object, List<Excel.Range>> valueDict,
        List<string> matches,
        List<Excel.Range> matchCells,
        List<Excel.Range> diffCells)
        {
            // 记录已匹配的区域二单元格
            HashSet<Excel.Range> matchedRegion2Cells = new HashSet<Excel.Range>();

            for (int row = 1; row <= values.GetLength(0); row++)
            {
                for (int col = 1; col <= values.GetLength(1); col++)
                {
                    Excel.Range cell1 = range.Cells[row, col];
                    object rawValue = values[row, col];
                    object keyValue = rawValue == null ? null : Convert.ChangeType(rawValue, typeof(string));

                    if (valueDict.TryGetValue(keyValue, out List<Excel.Range> region2Cells))
                    {
                        foreach (Excel.Range cell2 in region2Cells)
                        {
                            // 记录配对：区域一单元格 + 区域二单元格
                            matchCells.Add(cell1);
                            matchCells.Add(cell2);
                            matches.Add(rawValue?.ToString());

                            // 标记区域二单元格已匹配
                            matchedRegion2Cells.Add(cell2);
                        }
                    }
                    else
                    {
                        diffCells.Add(cell1); // 区域一独有的单元格
                    }
                }
            }

            // 查找区域二独有的单元格（未匹配的）
            foreach (var pair in valueDict)
            {
                foreach (Excel.Range cell in pair.Value)
                {
                    if (!matchedRegion2Cells.Contains(cell))
                        diffCells.Add(cell); // 区域二独有的单元格
                }
            }
        }
        // 辅助方法：更新界面
        private void UpdateUI(HashSet<string> 区域一独有, HashSet<string> 区域二独有, List<string> 相同值)
        {
            区域一Text.Text = string.Join(Environment.NewLine, 区域一独有);
            区域二Text.Text = string.Join(Environment.NewLine, 区域二独有);
            相同项Text.Text = string.Join(Environment.NewLine, 相同值.Distinct()); // 如需相同值去重可添加.Distinct()

            不同项.Enabled = 导出不同项.Enabled = 区域一独有.Count > 0;
            相同项.Enabled = 导出相同项.Enabled = 相同值.Count > 0;
        }

        private void 退出_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}