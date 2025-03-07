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
using 插件.MyCode;


namespace 插件
{
    public partial class ThisAddIn
    {
        private Excel.Application excelApp;
        private Range _lastHighlightedRange;
        private Color _highlightColor = Color.LightBlue;
        private 聚光灯 _聚光灯;
       
        //private KeyboardNavigationHandler _keyboardHandler;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Excel.Window activeWindow = Globals.ThisAddIn.Application.ActiveWindow;
            int currentScrollRow = activeWindow.ScrollRow;

            StaticClass.ExcelApp = Globals.ThisAddIn.Application;
            Excel.Workbook workbook1 = StaticClass.ExcelApp.Workbooks.Open("C:\\Users\\Administrator\\Desktop\\工作簿12.xlsx");
            _聚光灯 = new 聚光灯(this.Application);

            excelApp = this.Application;
            // 启动焦点窗口监控
            FocusMonitor.Start();
            
            // 订阅鼠标滚轮事件
            MouseHook.MouseWheelScrolled += OnMouseWheelScrolled;
            //MouseHook.MouseMiddleButtonClicked += OnMouseMiddleButtonClicked;
            //   mouseHook = new MouseScrollHook(this.Application);
            // mouseHook.StartHook();
            MessageBox.Show("Event subscriptions completed!");
        }
        // 初始化控制台
   
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
         
          
            // 停止焦点窗口监控
            FocusMonitor.Stop();

            // 停止鼠标钩子
            MouseHook.Stop();

            // 取消订阅事件
            MouseHook.MouseWheelScrolled -= OnMouseWheelScrolled;

          
           // MouseHook.MouseMiddleButtonClicked -= OnMouseMiddleButtonClicked;

            _聚光灯?.UnsubscribeEvents();
        }
        // 鼠标滚轮事件处理
        private void OnMouseWheelScrolled(int lines)
        {
            // 输出滚动的行数
           Console.WriteLine($"Mouse wheel scrolled {lines} lines!");

            // 在这里执行你的逻辑，例如滚动 Excel 单元格
            //ScrollExcelCells(lines);
        }


        // 鼠标中键事件处理
        private void OnMouseMiddleButtonClicked()
        {
            // 在这里执行你的逻辑
            //MessageBox.Show("Mouse middle button clicked!");
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