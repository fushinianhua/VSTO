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
            textBox5.Text = "0";
            textBox6.Text = "无";
            textBox7.Text = "重";
            progressBar1.Value = 100;
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

                string SourceKeyCol = Code1.StrtoW(Col1.Text);//源文件的key列
                string SourceValueCol = Code1.StrtoW(Col2.Text);//源文件的值列
                string ThisKeyCol = Code1.StrtoW(Col3.Text);//目标文件的key列
                string ThisValueCol = Code1.StrtoW(Col4.Text);//模板文件的值列

                Worksheet SourceSheet = (Worksheet)WKs[item].Worksheets[item2];//源文件
                Worksheet ThisSheet = (Worksheet)WKs[item3].Worksheets[item4];//目标文件

                long Source_rows = SourceSheet.Range[$"{SourceKeyCol}{SourceSheet.Rows.Count}"].End[XlDirection.xlUp].Row;//源文件的最后一行
                long This_rows = ThisSheet.Range[$"{ThisKeyCol}{ThisSheet.Rows.Count}"].End[XlDirection.xlUp].Row;//目标文件的最后一行

                Range S_Key = SourceSheet.Range[SourceKeyCol + "1"].Resize[Source_rows];
                Range S_Value = SourceSheet.Range[SourceValueCol + "1"].Resize[Source_rows];
                Range T_Key = ThisSheet.Range[ThisKeyCol + "1"].Resize[This_rows];
                Range T_Value = ThisSheet.Range[ThisValueCol + "1"].Resize[This_rows]; // 目标值列
                Dictionary<string, string> keyValues = 获取数据(S_Key, S_Value, Source_rows);//获取单元格数据
                List<string> 重复项 = new List<string>();
                int 重复数量 = 0;
                if (keyValues.Count > 0)
                {
                    for (int i = 2; i <= This_rows; i++)
                    {
                        Range rng = T_Key.Rows[i];//key列
                        string kry = rng.Value2?.ToString();//
                        Range rng2 = T_Value.Rows[i];//值列
                        try
                        {
                            if (string.IsNullOrEmpty(kry)) continue;
                            if (keyValues.ContainsKey(kry))
                            {
                                string newValue = keyValues[kry];//取到值
                                string currentValue = rng2.Value2?.ToString();
                                if (string.IsNullOrEmpty(currentValue))
                                {
                                    rng2.Value2 = newValue;
                                }
                                if (checkBox2.Checked)
                                {
                                    if (重复项.Contains(kry))
                                    {
                                        重复数量++;
                                        // rng2.Value2 = "重";
                                    }
                                    else
                                    {
                                        重复项.Add(kry);
                                    }
                                }

                            }
                            else
                            {
                                if (checkBox1.Checked)
                                {
                                    string newText = textBox6.Text;
                                    string currentValue = rng2.Value2?.ToString();
                                    if (currentValue != newText)
                                    {
                                        rng2.Value2 = newText;
                                    }
                                }
                            }
                        }
                        finally
                        {
                            // 释放 COM 对象
                            Marshal.ReleaseComObject(rng);
                            Marshal.ReleaseComObject(rng2);
                        }
                    }
                }
                
                MessageBox.Show($"共有重复数量：{重复数量}");
                TimeSpan timeSpan = DateTime.Now.Subtract(t0);
                double totalSeconds = timeSpan.TotalSeconds;
                textBox1.Text = totalSeconds + "秒";
                if (checkBox3.Checked)
                {
                    this.Close();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("发生错误: " + ex.Message);
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
                    Col4.Items.Add($"{ColNum + 1}.空白尾列");
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

                DialogResult result = MessageBox.Show("填充列已有数据,确认覆盖写入结果,\r,点击继续，中断操作点击取消。", "是否继续", MessageBoxButtons.OKCancel);
                if (result == DialogResult.OK)
                {
                    Run();
                }
                else { return; }

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
    }
}