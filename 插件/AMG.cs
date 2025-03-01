using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using 插件.MyCode;
using 插件.MyForm;

namespace 插件
{
    public partial class AMG
    {
        private void AMG_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.Form form = new 查询();
            if (form != null)
            {
                form.ShowDialog();
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            System.Windows.Forms.Form form = new 数据对比();
            if (form != null)
            {
                form.ShowDialog();
            }
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            Form form = new 聚光灯设置();
            form.ShowDialog();
        }
    }
}