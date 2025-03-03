using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using 插件.MyForm;
using System.Drawing;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;

using System.Runtime.InteropServices;
using ExcelMouseScrollVSTO;

namespace 插件
{
    public partial class ThisAddIn
    {
        private Excel.Application excelApp;
        private Range _lastHighlightedRange;
        private Color _highlightColor = Color.LightBlue;
        private 聚光灯 _聚光灯;
        private MouseScrollHook mouseHook;
        //private KeyboardNavigationHandler _keyboardHandler;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            StaticClass.ExcelApp = Globals.ThisAddIn.Application;
            //Excel.Workbook workbook1 = StaticClass.ExcelApp.Workbooks.Open("C:\\Users\\辛鹏\\Desktop\\工作簿12.xlsx");
            _聚光灯 = new 聚光灯(this.Application);

            excelApp = this.Application;
            //mouseHook = new MouseScrollHook(this.Application);
            //mouseHook.StartHook();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (mouseHook != null)
            {
                mouseHook.StopHook();
            }
            _聚光灯?.UnsubscribeEvents();
        }

        #region VSTO 生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion VSTO 生成的代码
    }
}