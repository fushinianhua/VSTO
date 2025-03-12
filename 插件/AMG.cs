using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
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
        public static Form 查询form = null;
        public static Form 对比form = null;
        public static Form 拆分form = null;
        public static Form 聚光灯form = null;
        private void AMG_Load(object sender, RibbonUIEventArgs e)
        {
            开光状态 = Settings.Default.聚光灯开关状态;
            StaticClass.聚光开关状态 = 开光状态;
            Setiamge(开光状态);
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {

                if (查询form == null)
                {
                    查询form = new 查询();
                }
                查询form.Show();
            }
            catch (Exception)
            {
            }
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (对比form == null)
                {
                    对比form = new 数据对比();
                }
                对比form.Show();
            }
            catch (Exception)
            {
            }
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (聚光灯form == null)
                {
                    聚光灯form = new 聚光灯设置();
                }
                聚光灯form.Show();

            }
            catch (Exception)
            {
            }
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

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                if (拆分form == null)
                {
                    拆分form = new 拆分工作表();
                }
                拆分form.Show();
            }
            catch (Exception)
            {
            }

        }

        private void button5_Click(object sender, RibbonControlEventArgs e)//升序
        {
            //ExecuteCommandSafely("SortAscendingExcel");
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("SortAscendingExcel");
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)//降序
        {
            // ExecuteCommandSafely("SortDescendingExcel");
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("SortDescendingExcel");
        }

        private void 重新应用_Click(object sender, RibbonControlEventArgs e)
        {
           // ExecuteCommandSafely("FilterReapply");
        }
        private bool 是否筛选=false;
        private void 筛选_Click(object sender, RibbonControlEventArgs e)
        {
            if (是否筛选)
            {
                清除();
                return;
            }
            是否筛选 = true;
            // 获取当前工作表的UsedRange
            Range usedRange =(Range) Globals.ThisAddIn.Application.Selection;
            usedRange.AutoFilter(Field: 1, Criteria1: Type.Missing, Operator: XlAutoFilterOperator.xlAnd, Criteria2: Type.Missing, VisibleDropDown: true);
        }

        private void 清除筛选_Click(object sender, RibbonControlEventArgs e)
        {
            if (!是否筛选) return;

            清除();
            // worksheet.AutoFilterMode = false;

        }
        private void 清除()
        {
            try
            {
                Microsoft.Office.Interop.Excel.Workbook workbook = Globals.ThisAddIn.Application.ActiveWorkbook;
                Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.ActiveSheet;
                worksheet.AutoFilterMode = false;

                是否筛选 = false;
            }
            catch (Exception)
            {

                throw;
            }
           
        }

        private void 自定义排序_Click(object sender, RibbonControlEventArgs e)
        {

        }
    }
}