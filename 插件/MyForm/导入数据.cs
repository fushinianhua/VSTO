using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Bson;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using 插件.Properties;
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
        private Application excelapp;
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

        private void 导入数据_Load(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(数据导入地址))
                {
                    PathText.Text = 数据导入地址 = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                }
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
                excelapp = new Application { Visible = false };
                选择工作薄 = excelapp.Workbooks.Open(sourceFilePath);

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
                // 拿到选择的导入列表头
                List<string> selectedItems = new List<string>();
                for (int i = 0; i < CheckList.Items.Count; i++)
                {
                    if (CheckList.GetItemChecked(i))
                    {
                        selectedItems.Add(CheckList.Items[i].ToString());
                    }
                }

                foreach (string selectedItem in selectedItems)
                {
                    bool isExist = false;
                    foreach (DataTypeInfo sourceItem in 列表数据)
                    {
                        if (sourceItem.Keywords.Contains(selectedItem))
                        {
                            isExist = true;
                            break;
                        }
                    }

                    if (!isExist)
                    {
                        MessageBox.Show("列表头不存在对应数据,请修改");
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
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
            Range rng = 选择工作表.UsedRange;
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
            if (excelapp != null)
            {
                excelapp.Quit();
                Marshal.ReleaseComObject(excelapp);
                excelapp = null;
            }
        }
        /// <summary>
        /// 读取导入列表头对应数据
        /// </summary>
        private void LoadConfig()
        {
            try
            {
                Assembly assembly = Assembly.GetExecutingAssembly();
                string resourceName = "插件.Resources.ColName.json";
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        MessageBox.Show("未找到配置文件资源。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                    using (StreamReader reader = new StreamReader(stream))
                    {
                        string json = reader.ReadToEnd();
                        using (JsonTextReader jsonReader = new JsonTextReader(new StringReader(json)))
                        {
                            JsonSerializer serializer = new JsonSerializer();
                            列表数据 = serializer.Deserialize<List<DataTypeInfo>>(jsonReader);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"读取配置文件时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }


}