using Microsoft.Office.Interop.Excel;


using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using 插件.Properties;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;

namespace 插件.MyForm
{
    public partial class 导入数据 : Form
    {
        public 导入数据()
        {
            InitializeComponent();
        }

        private string 数据导入地址 = Settings.Default.数据导入地址;
        private System.Data.DataTable dataTable = new System.Data.DataTable(); // 用于存储所有选中的列数据

        private void 导入数据_FormClosed(object sender, FormClosedEventArgs e)
        {
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
                currentSheet = Globals.ThisAddIn.Application.ActiveSheet;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载初始路径时发生错误：{ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
          
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
           
        }


        private void button2_Click(object sender, EventArgs e)
        {
            
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