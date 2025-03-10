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
        private readonly Color _highlightColor = Color.LightBlue;
        private 聚光灯 _聚光灯;     
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // 初始化Excel应用引用
            excelApp = this.Application;
            StaticClass.ExcelApp = excelApp;
            try
            {
               
                _聚光灯 = new 聚光灯(this.Application);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"初始化失败: {ex.Message}");
            }
            
        }
        // 初始化控制台  
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {       
            // 释放资源
            _聚光灯?.UnsubscribeEvents();
            if (_lastHighlightedRange != null)
            {
                Marshal.ReleaseComObject(_lastHighlightedRange);
            } 
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