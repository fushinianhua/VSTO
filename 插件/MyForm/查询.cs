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

namespace 插件.MyCode
{
    public partial class 查询 : Form
    {
        private Workbooks WKs = null;

        private string item = null;
        private string item2 = null;
        private string item3 = null;
        private string item4 = null;

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
            checkBox1.Checked = true;
            textBox5.Text = "2";
            textBox6.Text = "无";
            textBox7.Text = "重";
            progressBar1.PerformStep();
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

                string SourceKeyCol = Code1.StrtoW(Col1.Text);
                string SourceValueCol = Code1.StrtoW(Col2.Text);
                string ThisKeyCol = Code1.StrtoW(Col3.Text);
                string ThisValueCol = Code1.StrtoW(Col4.Text);

                Worksheet SourceSheet = (Worksheet)WKs[item].Worksheets[item2];
                Worksheet ThisSheet = (Worksheet)WKs[item3].Worksheets[item4];

                long Source_rows = SourceSheet.UsedRange.Rows.Count;
                long This_rows = ThisSheet.UsedRange.Rows.Count;

                Range S_Key = SourceSheet.Range[SourceKeyCol + "1"].Resize[Source_rows];
                Range S_Value = SourceSheet.Range[SourceValueCol + "1"].Resize[Source_rows];
                Range T_Key = ThisSheet.Range[ThisKeyCol + "1"].Resize[This_rows];
                Range T_Value = ThisSheet.Range[ThisValueCol + "1"].Resize[This_rows]; // 目标值列

                Dictionary<string, object> MapDict = BuildMapDict(SourceSheet, S_Key, S_Value, Source_rows);

                progressBar1.Maximum = (int)(This_rows - writenum + 1);

                for (long n = writenum; n <= This_rows; n++)
                {
                    string key = T_Key[n, 1].Value2?.ToString();
                    if (string.IsNullOrEmpty(key))
                    {
                        continue; // 如果键为空，跳过
                    }

                    // 检查目标单元格是否已经有数据
                    if (T_Value[n, 1].Value2 == null || string.IsNullOrEmpty(T_Value[n, 1].Value2.ToString()))
                    {
                        if (MapDict.ContainsKey(key))
                        {
                            T_Value[n, 1].Value2 = MapDict[key]; // 写入新值
                        }
                        else
                        {
                            T_Value[n, 1].Value2 = textBox6.Text; // 写入默认值
                        }
                    }

                    progressBar1.Value = (int)(n - writenum + 1);
                }

                TimeSpan timeSpan = DateTime.Now.Subtract(t0);
                double totalSeconds = timeSpan.TotalSeconds;
                textBox1.Text = totalSeconds + "秒";

                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("发生错误: " + ex.Message);
            }
        }

        private Dictionary<string, object> BuildMapDict(Worksheet SourceSheet, Range S_Key, Range S_Value, long Source_rows)
        {
            Dictionary<string, object> MapDict = new Dictionary<string, object>();
            for (long n = 1; n <= Source_rows; n++)
            {
                string key = S_Key[n, 1].Value2?.ToString();
                if (string.IsNullOrEmpty(key)) continue;

                if (!MapDict.ContainsKey(key))
                {
                    MapDict.Add(key, S_Value[n, 1].Value2);
                }
                else if (checkBox2.Checked)
                {
                    MapDict[key] = MapDict[key] + textBox7.Text + S_Value[n, 1].Value2;
                }
            }
            return MapDict;
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
                if (Col4.Text != "")
                {
                    string ColStr = Code1.StrtoW(Col4.Text);
                    Worksheet WS = (Worksheet)WKs[item].Worksheets[item2];
                    Range range = WS.Range[ColStr + ":" + ColStr];
                    int count = (int)StaticClass.ExcelApp.WorksheetFunction.CountA(range);
                    Tip4.Text = $"{ColStr}列:  {(double)count / 10000:0.000}万";
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
                item2 = comboBox2.SelectedItem.ToString();
                if (item != "")
                {
                    Col1.Items.Clear();
                    Col2.Items.Clear();
                    Worksheet WS = (Worksheet)WKs[item].Worksheets[item2];
                    long ColNum = WS.UsedRange.Columns.Count;
                    for (int i = 1; i < ColNum + 1; i++)
                    {
                        Range range = (Range)WS.Cells[1, i];
                        if (range.Value2 != "")
                        {
                            Col1.Items.Add($"{i}.{WS.Cells[1, i].Value2}");
                            Col2.Items.Add($"{i}.{WS.Cells[1, i].Value2}");
                        }
                        else
                        {
                            Col1.Items.Add($"{i}.null");
                            Col2.Items.Add($"{i}.null");
                        }
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
                item4 = comboBox4.SelectedItem.ToString();
                if (item != "")
                {
                    Col3.Items.Clear();
                    Col4.Items.Clear();
                    Worksheet WS = (Worksheet)WKs[item3].Worksheets[item4];
                    long ColNum = WS.UsedRange.Columns.Count;
                    for (int i = 1; i < ColNum + 1; i++)
                    {
                        Range range = (Range)WS.Cells[1, i];
                        if (range.Value2 != "")
                        {
                            Col3.Items.Add($"{i}.{WS.Cells[1, i].Value2}");
                            Col4.Items.Add($"{i}.{WS.Cells[1, i].Value2}");
                        }
                        else
                        {
                            Col3.Items.Add($"{i}.null");
                            Col4.Items.Add($"{i}.null");
                        }
                    }
                    Col4.Items.Add($"{ColNum}.空白尾列");
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
                if (Col1.Text == "" || Col2.Text == "" || Col3.Text == "" || Col4.Text == "")
                {
                    MessageBox.Show("请选择列");
                    return;
                }
                string ThisValueCol = Code1.StrtoW(Col4.Text);
                // 获取 writenum 的值
                int WriteNum = int.Parse(textBox5.Text);
                Worksheet WS = (Worksheet)WKs[item3].Worksheets[item4];
                Range range = (Range)WS.Range[ThisValueCol + ":" + ThisValueCol];
                int count = (int)StaticClass.ExcelApp.WorksheetFunction.CountA(range);
                Marshal.ReleaseComObject(range);
                if (count >= WriteNum)
                {
                    DialogResult result = MessageBox.Show("填充列已有数据,确认覆盖写入结果,\r,点击继续，中断操作点击取消。", "是否继续", MessageBoxButtons.OKCancel);
                    if (result == DialogResult.Cancel)
                    {
                        Run();
                    }
                    else { return; }
                }
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
    }
}