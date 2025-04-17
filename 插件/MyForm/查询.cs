using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using 插件.MyForm;

using Microsoft.Office.Interop.Excel;

using System.Diagnostics;
using System.IO;
using Newtonsoft.Json;
using static 插件.MyForm.StaticClass;
using System.ComponentModel.DataAnnotations;
using System.Management;
using System.Collections;

namespace 插件.MyCode
{
    public partial class 查询 : Form
    {
        private readonly Workbooks WKs = null;

        private string item = null;
        private string item2 = null;
        private string item3 = null;
        private string item4 = null;

        public bool IsChanged = false;

        private List<DataTypeInfo> 列表数据 = new List<DataTypeInfo>();

        public 查询()
        {
            InitializeComponent();
            WKs = StaticClass.ExcelApp.Workbooks;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            foreach (Workbook wk in WKs)
            {
                try
                {
                    comboBox1.Items.Add(wk.Name);
                    comboBox3.Items.Add(wk.Name);
                }
                finally
                {
                    // 释放 COM 对象
                    Marshal.ReleaseComObject(wk);
                }
            }
            if (comboBox1.Items.Count > 0)
            {
                comboBox1.SelectedIndex = 0;
                comboBox3.SelectedIndex = 0;
            }
            textBox5.Text = "0";
        }

        private void Run()
        {
            Worksheet 导入文件 = null;
            Worksheet 目标文件 = null;
            try
            {
                // 原始代码保持不变
                string item = comboBox1.SelectedItem?.ToString();
                string item2 = comboBox2.SelectedItem?.ToString();
                string item3 = comboBox3.SelectedItem?.ToString();
                string item4 = comboBox4.SelectedItem?.ToString();
                int 导入列索引 = Col1.SelectedIndex + 1;
                int 写入列索引 = Col3.SelectedIndex + 1;

                if (item == null || item2 == null || item3 == null || item4 == null)
                {
                    MessageBox.Show("请从所有下拉框中选择一个选项。");
                    return;
                }

                if (!long.TryParse(textBox5.Text, out long writenum))
                {
                    MessageBox.Show("textBox5 中的输入无效，请输入有效的数字。");
                    return;
                }
                导入文件 = (Worksheet)WKs[item].Worksheets[item2];
                目标文件 = (Worksheet)WKs[item3].Worksheets[item4];

                Range targetEndRange = 目标文件.Cells[目标文件.Rows.Count, 写入列索引];
                int targetEndRow = targetEndRange.End[XlDirection.xlUp].Row;
                Range targetRng = 目标文件.Range[目标文件.Cells[1, 写入列索引], 目标文件.Cells[targetEndRow, 写入列索引]];
                object[,] targetValues = targetRng.Value2;
                List<string> 目标列表 = 转换列表(targetValues);
                Marshal.ReleaseComObject(targetEndRange);
                Marshal.ReleaseComObject(targetRng);

                Range EndRange = 导入文件.Cells[导入文件.Rows.Count, 导入列索引];
                int EndRow = EndRange.End[XlDirection.xlUp].Row;
                Range rng = 导入文件.Range[导入文件.Cells[1, 导入列索引], 导入文件.Cells[EndRow, 导入列索引]];
                object[,] values = rng.Value2;
                List<string> 写入列表 = 转换列表(values);
                Marshal.ReleaseComObject(EndRange);
                Marshal.ReleaseComObject(rng);

                Range EndColumnRange = 导入文件.Cells[1, 导入文件.Columns.Count];
                int EndColumn = EndColumnRange.End[XlDirection.xlToLeft].Column;
                Range rngs = 导入文件.Range[导入文件.Cells[1, 1], 导入文件.Cells[EndRow, EndColumn]];
                object[,] 总数据 = rngs.Value2;
                Marshal.ReleaseComObject(EndColumnRange);
                Marshal.ReleaseComObject(rngs);

                List<string[]> 数据列表 = new List<string[]>();
                List<int> 映射列 = 映射列Dic.Keys.ToList();
                for (int i = 0; i < 写入列表.Count; i++)
                {
                    string[] data = new string[映射列Dic.Count];
                    for (int j = 0; j < 映射列Dic.Count; j++)
                    {
                        int a = 映射列[j];
                        data[j] = 总数据[i + 1, a]?.ToString();
                    }
                    数据列表.Add(data);
                }
                List<string> 已经写入值 = new List<string>();
                List<string> 重复项 = new List<string> { };

                for (int i = 0; i < 目标列表.Count; i++)
                {
                    string key = 目标列表[i];
                    if (string.IsNullOrEmpty(key)) continue;
                    if (写入列表.Contains(key))
                    {
                        // 修改点2：使用行号映射表查找真实行号
                        if (已经写入值.Contains(key))
                        {
                            重复项.Add(key);
                        }
                        已经写入值.Add(key);
                        for (int j = 0; j < 数据列表[i].Length; j++)
                        {
                            int 写入列 = 映射列Dic[映射列[j]];
                            // 修改点4：使用正确的目标行号
                            Range r = 目标文件.Cells[i + 1, 写入列];
                            string 单元格值 = r.Value2?.ToString();
                            if (string.IsNullOrEmpty(单元格值))
                            {
                                int row = 写入列表.IndexOf(key);
                                r.Value2 = 数据列表[row][j];
                            }
                            Marshal.ReleaseComObject(r); // 修改点5：及时释放对象
                        }
                    }
                }

                if (checkBox3.Checked)
                {
                    Close();
                }
                MessageBox.Show($"共有{重复项.Count}个重复项");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                Marshal.ReleaseComObject(目标文件);
                Marshal.ReleaseComObject(导入文件);
            }
        }

        private List<string> 转换列表(object[,] objects)
        {
            try
            {
                List<string> strings = new List<string>();
                if (objects != null)
                {
                    for (int i = 1; i <= objects.GetLength(0); i++)
                    {
                        for (int j = 1; j <= objects.GetLength(1); j++)
                        {
                            string value = objects[i, j]?.ToString();
                            if (!string.IsNullOrEmpty(value))
                            {
                                strings.Add(value);
                            }
                        }
                    }
                }

                return strings;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private Dictionary<string, string> 获取数据(Excel.Range KeyCol, Excel.Range ValueCol, long RowCount)
        {
            Dictionary<string, string> 数据 = new Dictionary<string, string>();
            for (int i = 2; i <= RowCount; i++)
            {
                Excel.Range rng = KeyCol.Rows[i];
                Excel.Range rng2 = ValueCol.Rows[i];
                string key = rng.Value2?.ToString();
                string value = rng2.Value2?.ToString();

                if (!string.IsNullOrEmpty(key) && !数据.ContainsKey(key))
                {
                    数据.Add(key, value);
                }

                // 释放 COM 对象
                Marshal.ReleaseComObject(rng);
                Marshal.ReleaseComObject(rng2);
            }
            return 数据;
        }

        private void Col1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (Col1.Text != "")
                {
                    string ColStr = Code1.StrtoW(Col1.Text);
                    Worksheet WS = (Worksheet)WKs[item].Worksheets[item2];
                    Range range = WS.Range[ColStr + ":" + ColStr];
                    int count = (int)StaticClass.ExcelApp.WorksheetFunction.CountA(range);
                    Tip1.Text = $"{ColStr}列:  {(double)count / 10000:0.000}万";
                }
            }
            catch
            {
            }
        }

        private void Col2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string ColText = Col2.Text;
                if (ColText == Col1.SelectedItem.ToString())
                {
                    Col2.SelectedIndexChanged -= Col2_SelectedIndexChanged;
                    Col2.SelectedIndex = -1;
                    Col2.SelectedIndexChanged += Col2_SelectedIndexChanged;
                    MessageBox.Show("列名重复");
                    return;
                }
                if (Col2.Text != "")
                {
                    string ColStr = Code1.StrtoW(Col2.Text);
                    Worksheet WS = (Worksheet)WKs[item].Worksheets[item2];
                    Range range = WS.Range[ColStr + ":" + ColStr];
                    int count = (int)StaticClass.ExcelApp.WorksheetFunction.CountA(range);
                    Tip2.Text = $"{ColStr}列:  {(double)count / 10000:0.000}万";
                }
            }
            catch
            {
            }
        }

        private void Col3_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (Col3.Text != "")
                {
                    string ColStr = Code1.StrtoW(Col3.Text);
                    Worksheet WS = (Worksheet)WKs[item3].Worksheets[item4];
                    Range range = WS.Range[ColStr + ":" + ColStr];
                    int count = (int)StaticClass.ExcelApp.WorksheetFunction.CountA(range);
                    Tip3.Text = $"{ColStr}列:  {(double)count / 10000:0.000}万";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void Col4_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string ColText = Col4.Text;
                if (ColText == Col3.SelectedItem.ToString())
                {
                    Col4.SelectedIndexChanged -= Col4_SelectedIndexChanged;
                    Col4.SelectedIndex = -1;
                    Col4.SelectedIndexChanged += Col4_SelectedIndexChanged;
                    MessageBox.Show("列名重复");
                    return;
                }
                if (Col4.Text != "")
                {
                    string ColStr = Code1.StrtoW(Col4.Text);
                    Worksheet WS = (Worksheet)WKs[item3].Worksheets[item4];
                    Range range = WS.Range[ColStr + ":" + ColStr];
                    int count = (int)StaticClass.ExcelApp.WorksheetFunction.CountA(range);
                    Tip4.Text = $"{ColStr}列:  {(double)count / 10000:0.000}万";
                }
            }
            catch (Exception)
            {
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                item = comboBox1.SelectedItem.ToString();
                if (item != "")
                {
                    Workbook workbook = WKs[item];
                    listBox1.Items.Clear();
                    上一次写入列 = -1;
                    空白列 = 1;
                    comboBox2.Text = "";
                    comboBox2.Items.Clear();
                    Col1.Text = "";
                    Col2.Text = "";
                    Col1.Items.Clear();
                    Col2.Items.Clear();
                    foreach (Worksheet ws in workbook.Worksheets)
                    {
                        comboBox2.Items.Add(ws.Name);
                    }

                    comboBox2.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                item2 = comboBox2.SelectedItem.ToString();
                if (item != "")
                {
                    listBox1.Items.Clear();
                    上一次写入列 = -1;
                    空白列 = 1;
                    Col1.Text = "";
                    Col2.Text = "";
                    Col1.Items.Clear();
                    Col2.Items.Clear();
                    Worksheet WS = (Worksheet)WKs[item].Worksheets[item2];
                    long ColNum = WS.Cells[1, WS.Columns.Count].End[XlDirection.xlToLeft].Column;

                    for (int i = 1; i < ColNum + 1; i++)
                    {
                        Range range = (Range)WS.Cells[1, i];
                        if (range.Value2 != "")
                        {
                            Col1.Items.Add($"{i}.{WS.Cells[1, i].Value2}");
                            Col2.Items.Add($"{i}.{WS.Cells[1, i].Value2}");
                        }
                    }
                    if (Col1.Items.Count > 0)
                    {
                        Col1.SelectedIndex = 0;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                item3 = comboBox3.SelectedItem.ToString();
                if (item != "")
                {
                    Workbook workbook = WKs[item3];
                    listBox1.Items.Clear();
                    上一次写入列 = -1;
                    空白列 = 1;
                    comboBox4.Items.Clear();
                    comboBox4.Text = "";
                    Col3.Text = "";
                    Col4.Text = "";
                    Col3.Items.Clear();
                    Col4.Items.Clear();
                    foreach (Worksheet ws in workbook.Worksheets)
                    {
                        comboBox4.Items.Add(ws.Name);
                    }
                    comboBox4.SelectedIndex = 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                item4 = comboBox4.SelectedItem.ToString();
                if (item != "")
                {
                    listBox1.Items.Clear();
                    上一次写入列 = -1;
                    空白列 = 1;
                    Col3.Text = "";
                    Col4.Text = "";
                    Col3.Items.Clear();
                    Col4.Items.Clear();
                    Worksheet WS = (Worksheet)WKs[item3].Worksheets[item4];
                    long ColNum = WS.Cells[1, WS.Columns.Count].End[XlDirection.xlToLeft].Column;
                    for (int i = 1; i < ColNum + 1; i++)
                    {
                        Range range = (Range)WS.Cells[1, i];
                        if (range.Value2 != "" && range.Value2 != null)
                        {
                            Col3.Items.Add($"{i}.{WS.Cells[1, i].Value2}");
                            Col4.Items.Add($"{i}.{WS.Cells[1, i].Value2}");
                        }
                    }
                    if (Col3.Items.Count > 0)
                    {
                        Col3.SelectedIndex = 0;
                    }
                    // Col4.Items.Add($"{ColNum + 1}.空白尾列");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private Dictionary<int, int> 映射列Dic = new Dictionary<int, int>();
        private int 上一次写入列 = -1;

        private void 添加项目_Click(object sender, EventArgs e)
        {
            try
            {
                string col1text = Col1.SelectedItem.ToString();
                string col2text = Col2.SelectedItem.ToString();
                string col3text = Col3.SelectedItem.ToString();
                string col4text = Col4.SelectedItem.ToString();

                if (col1text == col2text || col3text == col4text)
                {
                    MessageBox.Show("请选择与索引列不相同的列");
                    return;
                }
                int 导入列 = Col2.SelectedIndex + 1;
                int 写入列 = Col4.SelectedIndex + 1;
                if (映射列Dic.ContainsKey(导入列))
                {
                    MessageBox.Show("该列已经被选择过");
                    return;
                }
                if (写入列 == 上一次写入列)
                {
                    MessageBox.Show("该列已经被选择过");
                    return;
                }
                映射列Dic.Add(导入列, 写入列);
                上一次写入列 = Col4.SelectedIndex + 1;
                listBox1.Items.Add($"{Col2.SelectedItem}-- {Col4.SelectedItem}");
                Col2.SelectedIndex = -1;
                Col4.SelectedIndex = -1;
            }
            catch (Exception)
            {
            }
        }

        //开始按钮
        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (Col1.Text == "" || Col3.Text == "")
                {
                    MessageBox.Show("请选择列");
                    return;
                }
                Run();
            }
            catch
            {
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            try
            {
                listBox1.Items.Clear();
                上一次写入列 = -1;
                空白列 = 1;
                comboBox1.Items.Clear();
                comboBox2.Items.Clear();
                Col1.Items.Clear();
                Col2.Items.Clear();
                foreach (Workbook wk in WKs)
                {
                    try
                    {
                        comboBox1.Items.Add(wk.Name);
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(wk);
                    }
                }
                if (comboBox1.Items.Count > 0)
                {
                    comboBox1.SelectedIndex = 0;
                }
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
                listBox1.Items.Clear();
                上一次写入列 = -1;
                空白列 = 1;
                comboBox3.Items.Clear();
                comboBox4.Items.Clear();
                Col4.Items.Clear();
                Col3.Items.Clear();
                foreach (Workbook wk in WKs)
                {
                    try
                    {
                        comboBox3.Items.Add(wk.Name);
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(wk);
                    }
                }
                if (comboBox3.Items.Count > 0)
                {
                    comboBox3.SelectedIndex = 0;
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        private void 查询_FormClosed(object sender, FormClosedEventArgs e)
        {
            Globals.ThisAddIn.查询form = null;
        }

        private int 空白列 = 1;

        private void button1_Click(object sender, EventArgs e)
        {
            if (Col4.Items.Count > 0)
            {
                Col4.Items.Add("空白列" + 空白列);
                空白列++;
                Col4.SelectedIndex = Col4.Items.Count - 1;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            映射列Dic.Clear();
            listBox1.Items.Clear();
            上一次写入列 = 0;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int count = 映射列Dic.Count - 1;
            var key = 映射列Dic.Keys.ToList();
            映射列Dic.Remove(key[key.Count - 1]);
            var allitem = listBox1.Items;
            allitem.Remove(allitem.Count - 1);
            listBox1.Items.Clear();
            listBox1.Items.AddRange(allitem);
            上一次写入列 = 0;
        }
    }
}