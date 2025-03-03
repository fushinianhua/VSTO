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
using 插件.Properties;

namespace 插件
{
    public partial class AMG
    {
        private bool 开光状态;
        private void AMG_Load(object sender, RibbonUIEventArgs e)
        {
            开光状态 = Settings.Default.聚光灯开关状态;
            StaticClass.聚光开关状态 = 开光状态;
            Setiamge(开光状态);
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

        private void 聚光灯_Click(object sender, RibbonControlEventArgs e)
        {
            开光状态 = !开光状态;
            StaticClass.聚光开关状态 = 开光状态;
            Setiamge(开光状态);
            Settings.Default.聚光灯开关状态 = 开光状态;
            Settings.Default.Save();
        }
        private void Setiamge(bool value)
        {
            if (value)
            {
                聚光灯.Image = Resources.聚光灯开;
            }
            else
            {
                聚光灯.Image = Resources.聚光灯关;
            }
        }
    }
}