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

namespace 插件.MyForm
{
    public partial class 数据对比 : Form
    {
        public 数据对比()
        {
            InitializeComponent();
           
        }
        private Excel.Application excelapp;

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
            else
            { 清除标识.Enabled = true; }
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

        private Excel.Range 区域一 = null;
        private Excel.Range 区域二 = null;

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        private const uint SWP_SHOWWINDOW = 0x0040;
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            // 获取 Excel 窗口句柄
            IntPtr excelHwnd = (IntPtr)excelapp.Hwnd;
            Form form = new 数据对比();
            form.Hide();
            //
            //// 使用 InputBox 方法提示用户选择单元格
            //object result = excelapp.InputBox(Prompt: "请选择单元格", Title: "选择单元格", Default: "", Type: 8 // Type 8 表示返回一个 Range 对象
            //);
            this.Hide();
            // 确保 Excel 窗口显示
            SetWindowPos(excelHwnd, IntPtr.Zero, 0, 0, 0, 0, SWP_SHOWWINDOW);
            //// 检查用户是否取消了选择
            //if (result != null)
            //{
            //    // 将结果转换为 Range 对象    

            //    if (result is Excel.Range selectedRange)
            //    {
            //        区域一 = selectedRange;
            //        string str = selectedRange.Address.ToString().Replace("$", "");
            //        区域1Box.Text = str;

            //        // 释放 COM 对象
            //        Marshal.ReleaseComObject(selectedRange);
            //    }
            //}
            Thread.Sleep(1000);
            form.Show();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            this.Hide();
            // 使用 InputBox 方法提示用户选择单元格
            object result = excelapp.InputBox(Prompt: "请选择单元格", Title: "选择单元格", Default: "", Type: 8 // Type 8 表示返回一个 Range 对象
            );

            // 检查用户是否取消了选择
            if (result != null)
            {
                // 将结果转换为 Range 对象

                if (result is Excel.Range selectedRange)
                {
                    区域二 = selectedRange;
                    string str = selectedRange.Address.ToString().Replace("$", "");
                    区域2Box.Text = str;

                    // 释放 COM 对象
                    Marshal.ReleaseComObject(selectedRange);
                }
            }
            this.Show();
        }

        private void 对比数据_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(区域1Box.Text))
                {
                    MessageBox.Show("区域一未选择");
                }
                if (string.IsNullOrEmpty(区域2Box.Text))
                {
                    MessageBox.Show("区域二未选择");
                }
                object[,] valuesArray1 = 区域一.Value2 as object[,];
                object[,] valuesArray2 = 区域二.Value2 as object[,];
                // 直接将单元格区域的值加载到数组中
             //   object[,] valuesArray1 = 区域一.Value2;
             //  object[,] valuesArray2 = 区域一.Value2;
                HashSet<object> list = new HashSet<object>();
                for (int i = 1; i <= valuesArray2.GetLength(0); i++)
                {
                    list.Add(valuesArray2[i, 1]);
                }
                // 存储相同的值和区域一中不同的值
                List<object> sameValues = new List<object>();
                List<object> differentValues1 = new List<object>();

                // 遍历区域一的值
                for (int i = 1; i <= valuesArray1.GetLength(0); i++)
                {
                    object value = valuesArray1[i, 1];
                    if (list.Contains(value))
                    {
                        sameValues.Add(value);
                    }
                    else
                    {
                        differentValues1.Add(value);
                    }
                }
                // 创建 HashSet 存储区域一的值，用于快速查找
                HashSet<object> set1 = new HashSet<object>(sameValues.Concat(differentValues1));

                // 存储区域二中不同的值
                List<object> differentValues2 = new List<object>();
                for (int i = 1; i <= valuesArray2.GetLength(0); i++)
                {
                    object value = valuesArray2[i, 1];
                    if (!set1.Contains(value))
                    {
                        differentValues2.Add(value);
                    }
                }
                区域一Text.AppendText(differentValues1.ToArray().ToString());
                区域二Text.AppendText(differentValues2.ToArray().ToString());
                相同项Text.AppendText(sameValues.ToArray().ToString());

            }
            catch (Exception)
            {

                throw;
            }


        }

        private void 退出_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}