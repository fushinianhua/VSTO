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

namespace 插件.MyForm
{
    public partial class 数据对比 : Form
    {
        public 数据对比()
        {
            InitializeComponent();
            KeyPreview = true;
            KeyDown += Form2_KeyDown;
        }

        private bool _Is相同项标识 = false;
        private bool _Is不同项标识 = false;

        private bool Is不同项标识
        {
            set
            {
                if (_Is不同项标识 == false)
                {
                    butt();
                }
                _Is不同项标识 = value;
            }
        }

        private bool Is相同项标识
        {
            set
            {
                if (_Is不同项标识 == false)
                {
                    butt();
                }
                _Is不同项标识 = value;
            }
        }

        private void butt()
        {
            if (_Is相同项标识 = _Is相同项标识 = false)
            { 清除标识.Enabled = false; }
        }

        private void 相同项_Click(object sender, EventArgs e)
        {
            Is相同项标识 = true;
            Button button = (Button)sender;
            MessageBox.Show(button.Text);
        }

        private void 不同项_Click(object sender, EventArgs e)
        {
            Is相同项标识 = true;
            Button button = (Button)sender;
            MessageBox.Show(button.Text);
        }

        private void 清除标识_Click(object sender, EventArgs e)
        {
            _Is不同项标识 = _Is相同项标识 = false;
            Button button = (Button)sender;
            MessageBox.Show(button.Text);
        }

        private void 导出相同项_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            MessageBox.Show(button.Text);
        }

        private void 导出不同项_Click(object sender, EventArgs e)
        {
            Button button = (Button)sender;
            MessageBox.Show(button.Text);
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            相同项.Enabled = false;
            不同项.Enabled = false;
            清除标识.Enabled = false;
            导出不同项.Enabled = false;
            导出相同项.Enabled = false;
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

        private Excel.Range 区域一 = null;
        private Excel.Range 区域二 = null;

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            // 使用 InputBox 方法提示用户选择单元格
            object result = StaticClass.ExcelApp.InputBox(Prompt: "请选择单元格", Title: "选择单元格", Default: "", Type: 8 // Type 8 表示返回一个 Range 对象
            );

            // 检查用户是否取消了选择
            if (result != null)
            {
                // 将结果转换为 Range 对象

                if (result is Excel.Range selectedRange)
                {
                    区域一 = selectedRange;
                    string str = selectedRange.Address.ToString().Replace("$", "");
                    textBox1.Text = str;

                    // 释放 COM 对象
                    Marshal.ReleaseComObject(selectedRange);
                }
            }
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            // 使用 InputBox 方法提示用户选择单元格
            object result = StaticClass.ExcelApp.InputBox(Prompt: "请选择单元格", Title: "选择单元格", Default: "", Type: 8 // Type 8 表示返回一个 Range 对象
            );

            // 检查用户是否取消了选择
            if (result != null)
            {
                // 将结果转换为 Range 对象

                if (result is Excel.Range selectedRange)
                {
                    区域二 = selectedRange;
                    string str = selectedRange.Address.ToString().Replace("$", "");
                    textBox2.Text = str;

                    // 释放 COM 对象
                    Marshal.ReleaseComObject(selectedRange);
                }
            }
        }

        private void 对比数据_Click(object sender, EventArgs e)
        {
        }

        private void 退出_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}