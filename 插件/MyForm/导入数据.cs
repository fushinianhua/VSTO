
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Bson;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using 插件.Properties;
using static System.Net.WebRequestMethods;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static 插件.MyForm.StaticClass;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace 插件.MyForm
{
    public partial class 导入数据 : Form
    {
        public 导入数据()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 源文本数据地址
        /// </summary>
        private string sourceFilePath;
        private string 数据导入地址 = Settings.Default.数据导入地址;
        private Workbook 选择工作薄;
        private Worksheet 选择工作表;
        private object[,] 数据;
        private readonly List<string> 工作表名字 = new List<string>();
        private readonly List<string> 导入列表头 = new List<string>();
        private List<DataTypeInfo> 列表数据 = new List<DataTypeInfo>();
        // 新增：用于保存 A2 - H2 单元格格式的列表
        private Dictionary<int, string> RangeFormat = new Dictionary<int, string>();
        private void 导入数据_FormClosed(object sender, FormClosedEventArgs e)
        {
            ReleaseExcelObjects();
            Globals.ThisAddIn.导入form = null;

        }
        private Worksheet targetSheet;
        private void 导入数据_Load(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(数据导入地址))
                {
                    PathText.Text = 数据导入地址 = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                }
                targetSheet = Globals.ThisAddIn.Application.ActiveSheet;
                // 为按钮添加提示信息
                toolTip1.SetToolTip(button1, "点击选择需要导入的文件");
                toolTip1.SetToolTip(button3, "点击修改近义匹配列表");
                toolTip1.SetToolTip(button2, "点击开始导入数据");
                toolTip1.SetToolTip(comboBox1, "可以选择sheet,默认为第一个sheet");
                LoadConfig();
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
               
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.InitialDirectory = 数据导入地址;
                    openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                    openFileDialog.Title = "选择Excel文件";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        sourceFilePath = openFileDialog.FileName;
                        PathText.Text = sourceFilePath;
                        if (CheckList.Items.Count > 0)
                        {
                            CheckList.Items.Clear();
                        }
                        LoadHeadersFromSource();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"选择文件时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 加载表头
        private void LoadHeadersFromSource()
        {
            try
            {

                选择工作薄 = Globals.ThisAddIn.Application.Workbooks.Open(sourceFilePath);
                GetWorksheetNames();
                选择工作表 = 选择工作薄.Worksheets[1];
                LoadDataAndHeaders();

                if (导入列表头.Count > 0)
                {
                    PopulateCheckList();
                }

                comboBox1.SelectedIndexChanged -= comboBox1_SelectedIndexChanged;
                if (工作表名字.Count > 0)
                {
                    comboBox1.Items.AddRange(工作表名字.ToArray());
                    comboBox1.SelectedIndex = 0;
                }
                comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载表头时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                ReleaseExcelObjects();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string item = comboBox1.SelectedItem.ToString();
                if (工作表名字.Contains(item))
                {
                    RangeFormat.Clear();
                    导入列表头.Clear();
                    数据 = null;
                    CheckList.Items.Clear();

                    选择工作表 = 选择工作薄.Worksheets[item];
                    LoadDataAndHeaders();

                    if (导入列表头.Count > 0)
                    {
                        PopulateCheckList();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"切换工作表时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                // 获取选中的源列头
                var selectedHeaders = CheckList.CheckedItems.Cast<string>().ToList();
                if (selectedHeaders.Count == 0)
                {
                    MessageBox.Show("请至少选择一列进行导入");
                    return;
                }

                // 获取目标工作表

                var targetHeaders = GetTargetHeaders(targetSheet); // 获取目标表头

                // 建立列映射关系（源列索引 -> 目标列索引）
                var columnMapping = new Dictionary<int, int>();
                foreach (var srcHeader in selectedHeaders)
                {
                    // 查找近义词配置
                    DataTypeInfo dataType = 列表数据.FirstOrDefault(d => d.Keywords.Contains(srcHeader));
                    if (dataType == null)
                    {
                        MessageBox.Show($"未找到【{srcHeader}】的列配置");
                        continue;
                    }

                    // 在目标表中查找匹配列
                    var targetCol = FindTargetColumn(targetHeaders, dataType.Keywords);
                    if (targetCol == -1)
                    {
                        MessageBox.Show($"未找到【{srcHeader}】对应的目标列");
                        continue;
                    }

                    // 记录映射关系
                    int srcColIndex = 导入列表头.IndexOf(srcHeader) + 1; // Excel列从1开始
                    columnMapping.Add(srcColIndex, targetCol);
                }

                // 执行数据填充
                FillDataToTarget(columnMapping, targetSheet);

                MessageBox.Show($"成功导入 {columnMapping.Count} 列数据");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"导入失败：{ex.Message}");
            }
            finally
            {
                ReleaseExcelObjects();
            }
        }
        /// <summary>
        /// 在目标表中查找匹配列
        /// </summary>
        private int FindTargetColumn(Dictionary<int, string> targetHeaders, List<string> keywords)
        {
            foreach (var keyword in keywords)
            {
                var match = targetHeaders.FirstOrDefault(h =>
                    h.Value.Equals(keyword, StringComparison.OrdinalIgnoreCase));
                if (!match.Equals(default(KeyValuePair<int, string>)))
                {
                    return match.Key;
                }
            }
            return -1; // 未找到
        }

        /// <summary>
        /// 填充数据到目标列
        /// </summary>
        private void FillDataToTarget(Dictionary<int, int> columnMapping, Worksheet targetSheet)
        {
            int startRow = targetSheet.UsedRange.Rows.Count + 1;

            int maxRow = 数据.GetLength(0);

            foreach (var mapping in columnMapping)
            {
                int srcCol = mapping.Key;
                int targetCol = mapping.Value;

                // 准备数据数组（优化写入性能）
                object[,] dataArray = new object[maxRow - 1, 1]; // 行数从2开始
                for (int row = 2; row <= maxRow; row++)
                {
                    dataArray[row - 2, 0] = 数据[row, srcCol] ?? DBNull.Value;
                }

                // 批量写入数据
                Range targetRange = (Range)targetSheet.Range[
                    targetSheet.Cells[startRow, targetCol],
                    targetSheet.Cells[startRow + maxRow - 2, targetCol]
                ];
                // 先设置格式，再写入数据
                string format = RangeFormat.ContainsKey(srcCol) ? RangeFormat[srcCol] : "@";
                targetRange.NumberFormat = format;  // 先设置格式
                targetRange.Value2 = dataArray;     // 再写入数据
                // 强制刷新 Excel 应用程序
                targetSheet.Application.CalculateFull();
            }
        }


        /// <summary>
        /// 获取目标表头信息（列索引 -> 列名）
        /// </summary>
        private Dictionary<int, string> GetTargetHeaders(Worksheet targetSheet)
        {
            var headers = new Dictionary<int, string>();
            Range usedRange = targetSheet.UsedRange;
            Range headerRow = usedRange.Rows[1]; // 假设表头在第一行

            foreach (Range cell in headerRow.Cells)
            {
                if (cell.Value2 != null)
                {
                    headers[cell.Column] = cell.Value2.ToString().Trim();
                }
            }
            return headers;
        }

        private bool ValidateHeaders(List<string> headers)
        {
            foreach (var header in headers)
            {
                bool exists = 列表数据.Any(d => d.Keywords.Contains(header));
                if (!exists) return false;
            }
            return true;
        }


        /// <summary>
        /// 获取所有工作表名字
        /// </summary>
        private void GetWorksheetNames()
        {
            工作表名字.Clear();
            foreach (Worksheet ws in 选择工作薄.Worksheets)
            {
                工作表名字.Add(ws.Name);
            }
        }
        /// <summary>
        /// 获取sheet导入列表头
        /// </summary>
        private void LoadDataAndHeaders()
        {
            int row = 选择工作表.Cells[选择工作表.Rows.Count, 1].End[XlDirection.xlUp].Row;
            int col = 选择工作表.Cells[1, 选择工作表.Columns.Count].End[XlDirection.xlToLeft].Column;

            Range rng = 选择工作表.Range[选择工作表.Cells[1, 1], 选择工作表.Cells[row, col]];
            数据 = rng.Value2;
            导入列表头.Clear();
            if (数据 != null && 数据.GetLength(0) > 0 && 数据.GetLength(1) > 0)
            {
                for (int i = 1; i <= 数据.GetLength(1); i++)
                {
                    Range r = rng[2, i];
                    string format = r.NumberFormat;
                    RangeFormat.Add(i, format);
                    导入列表头.Add(数据[1, i].ToString());
                }
            }
        }
        /// <summary>
        /// 添加CheckList的item
        /// </summary>
        private void PopulateCheckList()
        {
            CheckList.Items.Clear();
            for (int i = 0; i < 导入列表头.Count; i++)
            {
                CheckList.Items.Add(导入列表头[i]);
                CheckList.SetItemChecked(i, true);
            }
        }
        /// <summary>
        /// 释放资源
        /// </summary>
        private void ReleaseExcelObjects()
        {
            if (选择工作表 != null)
            {
                Marshal.ReleaseComObject(选择工作表);
                选择工作表 = null;
            }
            if (选择工作薄 != null)
            {
                选择工作薄.Close(false);
                Marshal.ReleaseComObject(选择工作薄);
                选择工作薄 = null;
            }
        }
        /// <summary>
        /// 读取导入列表头对应数据
        /// </summary>
        private void LoadConfig()
        {
            string jsonFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "ColName.json");
            try
            {
                using (FileStream stream = new FileStream(jsonFilePath, FileMode.Open, FileAccess.Read))
                {
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        string json = reader.ReadToEnd();
                        using (JsonTextReader jsonReader = new JsonTextReader(new StringReader(json)))
                        {
                            JsonSerializer serializer = new JsonSerializer();
                            List<DataTypeInfo> listData = serializer.Deserialize<List<DataTypeInfo>>(jsonReader);
                            // 这里假设列表数据是类中的成员变量，将反序列化后的数据赋值给它
                            列表数据 = listData;
                        }
                    }
                }
            }
            catch (FileNotFoundException)
            {
                MessageBox.Show("配置文件未找到。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (JsonException ex)
            {
                MessageBox.Show($"解析JSON时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"发生其他错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                bool isChecked = checkBox1.Checked;
                for (int i = 0; i < CheckList.Items.Count; i++)
                {
                    CheckList.SetItemChecked(i, isChecked);
                }
                checkBox1.Text = isChecked ? "全部取消" : "全部选中";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"全选/全不选时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        bool IsChanged = false;
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                string jsonFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "ColName.json");

                Process.Start("notepad.exe", jsonFilePath);
                IsChanged = true;
                button4.Enabled = IsChanged;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (!IsChanged) return;
                LoadConfig();
                MessageBox.Show("数据更新成功");
                IsChanged = false;
                button4.Enabled = IsChanged;
            }
            catch (Exception)
            {

                throw;
            }
        }
    }


}