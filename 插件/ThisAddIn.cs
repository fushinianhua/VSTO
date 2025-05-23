﻿using System;
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
        public Form 查询form = null;
        public Form 对比form = null;
        public Form 拆分form = null;
        public Form 聚光灯form = null;
        public Form 导出form = null;
        public Form 导入form = null;

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

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // 释放资源
            _聚光灯?.UnsubscribeEvents();
            if (_lastHighlightedRange != null)
            {
                Marshal.ReleaseComObject(_lastHighlightedRange);
                _lastHighlightedRange = null;
            }

            if (excelApp != null)
            {
                Marshal.ReleaseComObject(excelApp);
                excelApp = null;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
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