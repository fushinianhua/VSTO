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
        private Workbooks WKs = null;

        private string item = null;
        private string item2 = null;
        private string item3 = null;
        private string item4 = null;

        public bool IsChanged = false;

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
            textBox6.Text = "无";
            textBox7.Text = "重";
            progressBar1.Value = 100;
            LoadConfig();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            progressBar1.PerformStep();
        }

        private void Run()
        {
            try
            {
                DateTime t0 = DateTime.Now; // 记录开始时间

                string item = comboBox1.SelectedItem?.ToString();
                string item2 = comboBox2.SelectedItem?.ToString();
                string item3 = comboBox3.SelectedItem?.ToString();
                string item4 = comboBox4.SelectedItem?.ToString();

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
                Worksheet 读取的文件 = (Worksheet)WKs[item].Worksheets[item2];//源文件
                Worksheet 写入的文件 = (Worksheet)WKs[item3].Worksheets[item4];//目标文件
                Range EndRange = 写入的文件.Cells[1, 写入的文件.Columns.Count];
                int EndColunm = EndRange.End[XlDirection.xlToLeft].Column;
                Dictionary<string, int> 写入的数据表头 = Enumerable.Range(0, CheckList2.Items.Count)
                                                        .Where(i => CheckList2.GetItemChecked(i))
                                                        .ToDictionary(i => CheckList2.Items[i].ToString(), i => i + 1);

                Dictionary<string, int> 读取的数据表头 = Enumerable.Range(0, CheckList1.Items.Count)
                                                     .ToDictionary(i => CheckList1.Items[i].ToString(), i => i + 1);
                DataTypeInfo dataType = 列表数据.FirstOrDefault(k => 写入的数据表头.Keys.Any(key => k.Keywords.Contains(key)));

                //if (dataType.Keywords.Count == dataType2.Keywords.Count)
                //{
                //    foreach (var key in dataType.Keywords)
                //    {
                //        if (dataType2.Keywords.Contains(key))
                //        {
                //            int a = 写入数据表头;
                //            int a=dataType.Keywords.IndexOf(key) + 1;
                //            columnMapping.Add(写入数据表头., dataType2.Keywords.IndexOf(key) + 1);
                //        }
                //        else
                //        {
                //            columnMapping.Add(dataType.Keywords.IndexOf(key) + 1, EndColunm + 1);
                //            EndColunm++;
                //        }
                //    }
                //}
                //for (int i = 0; i < columnMapping.Count; i++)
                //{
                //}
                //if (keyValues.Count > 0)
                //{
                //    for (int i = 2; i <= This_rows; i++)
                //    {
                //        Range rng = T_Key.Rows[i];//key列
                //        string kry = rng.Value2?.ToString();//
                //        Range rng2 = T_Value.Rows[i];//值列
                //        try
                //        {
                //            if (string.IsNullOrEmpty(kry)) continue;
                //            if (keyValues.ContainsKey(kry))
                //            {
                //                string newValue = keyValues[kry];//取到值
                //                string currentValue = rng2.Value2?.ToString();
                //                if (string.IsNullOrEmpty(currentValue))
                //                {
                //                    rng2.Value2 = newValue;
                //                }
                //                if (checkBox2.Checked)
                //                {
                //                    if (重复项.Contains(kry))
                //                    {
                //                        重复数量++;
                //                        // rng2.Value2 = "重";
                //                    }
                //                    else
                //                    {
                //                        重复项.Add(kry);
                //                    }
                //                }
                //            }
                //            else
                //            {
                //                if (checkBox1.Checked)
                //                {
                //                    string newText = textBox6.Text;
                //                    string currentValue = rng2.Value2?.ToString();
                //                    if (currentValue != newText)
                //                    {
                //                        rng2.Value2 = newText;
                //                    }
                //                }
                //            }
                //        }
                //        finally
                //        {
                //            // 释放 COM 对象
                //            Marshal.ReleaseComObject(rng);
                //            Marshal.ReleaseComObject(rng2);
                //        }
            }
            catch
            { }
            //}
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
                    需求空白列text.Text = $"{ColStr}列:  {(double)count / 10000:0.000}万";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
                    comboBox2.Items.Clear();
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
                if (CheckList1.Items.Count > 0)
                {
                    CheckList1.Items.Clear();
                }
                item2 = comboBox2.SelectedItem.ToString();
                if (item != "")
                {
                    Col1.Items.Clear();
                    Worksheet WS = (Worksheet)WKs[item].Worksheets[item2];
                    long ColNum = WS.Cells[1, WS.Columns.Count].End[XlDirection.xlToLeft].Column;
                    List<string> strings = new List<string>();
                    for (int i = 1; i < ColNum + 1; i++)
                    {
                        Range range = (Range)WS.Cells[1, i];
                        if (range.Value2 != "")
                        {
                            Col1.Items.Add($"{i}.{WS.Cells[1, i].Value2}");
                            //Col2.Items.Add($"{i}.{WS.Cells[1, i].Value2}");
                            strings.Add(WS.Cells[1, i].Value2);
                        }
                    }
                    int index = 0;
                    foreach (string str in strings)
                    {
                        CheckList1.Items.Add(str);
                        //CheckList1.SetItemChecked(index, true);
                        index++;
                    }

                    Tip2.Text = $"共选择了{CheckList1.CheckedItems.Count.ToString()}项数据";
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
                    comboBox4.Items.Clear();
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
                if (CheckList2.Items.Count > 0)
                {
                    CheckList2.Items.Clear();
                }
                item4 = comboBox4.SelectedItem.ToString();
                if (item != "")
                {
                    Col3.Items.Clear();
                    Worksheet WS = (Worksheet)WKs[item3].Worksheets[item4];
                    long ColNum = WS.Cells[1, WS.Columns.Count].End[XlDirection.xlToLeft].Column;
                    List<string> strings = new List<string>();
                    for (int i = 1; i < ColNum + 1; i++)
                    {
                        Range range = (Range)WS.Cells[1, i];
                        if (range.Value2 != "" && range.Value2 != null)
                        {
                            Col3.Items.Add($"{i}.{WS.Cells[1, i].Value2}");
                            strings.Add(WS.Cells[1, i].Value2);
                        }
                    }
                    int index = 0;
                    foreach (string str in strings)
                    {
                        CheckList2.Items.Add(str);
                        index++;
                    }
                    int.TryParse(需求空白列text.Text, out int count);
                    for (int i = 0; i < count; i++)
                    {
                        CheckList1.Items.Add("空白列");
                    }

                    Tip4.Text = $"共选择了{CheckList1.CheckedItems.Count}项数据";
                    // Col4.Items.Add($"{ColNum + 1}.空白尾列");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
            comboBox1.Items.Clear();
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

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            comboBox3.Items.Clear();
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

        private void 查询_FormClosed(object sender, FormClosedEventArgs e)
        {
            Globals.ThisAddIn.查询form = null;
        }

        private void CheckList1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Tip2.Text = $"共选择了{CheckList1.CheckedItems.Count.ToString()}项数据";
        }

        private void 需求空白列text_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int.TryParse(需求空白列text.Text, out int count);
                foreach (string str in CheckList1.Items)
                {
                    if (str == "空白列")
                    {
                        CheckList1.Items.Remove(str);
                    }
                }
                for (int i = 0; i < count; i++)
                {
                    CheckList1.Items.Add("空白列");
                }
            }
            catch (Exception)
            {
            }
        }

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

        private List<DataTypeInfo> 列表数据 = new List<DataTypeInfo>();

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
    }
}