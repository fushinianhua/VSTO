using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 插件.MyForm
{
    public partial class 导出数据 : Form
    {
        public 导出数据()
        {
            InitializeComponent();
            工作表 = StaticClass.ExcelApp.ActiveSheet;
        }
        private Worksheet 工作表 = null;
        List<string> 列名 = new List<string>();
        object[,] value = null;
        private void 导出数据_Load(object sender, EventArgs e)
        {

            if (工作表 != null)
            {
                Range r = (Range)工作表.Cells[1, 工作表.Columns.Count];//最后一列
                int col = r.End[XlDirection.xlToLeft].Column;
                Range c = (Range)工作表.Cells[工作表.Rows.Count,1 ];//最后一行
                int row = c.End[XlDirection.xlUp].Row;
                Range 表头单元格 = 工作表.Range[工作表.Cells[1,1],工作表.Cells[1,col]];
                Range 数据单元格 = 工作表.Range[工作表.Cells[1, 1], 工作表.Cells[row, col]];
                value = 数据单元格.Value2;
                for (int i = 1; i <= col; i++)
                {
                    CheckList.Items.Add(value[1,i]);
                }
                Marshal.ReleaseComObject(r);
                Marshal.ReleaseComObject(c);
                Marshal.ReleaseComObject(数据单元格);
                Marshal.ReleaseComObject(表头单元格);


            }
        }
        private void 导出数据_FormClosed(object sender, FormClosedEventArgs e)
        {
            Globals.ThisAddIn.导出form = null;
        }

        private void 文件导出_Click(object sender, EventArgs e)
        {
            if (CheckList.CheckedItems.Count > 0)
            {
                //执行选择的列作为数据

            }
            else
            {
                MessageBox.Show("未选择任何列的数据导出");
                return;
            }
        }
    }
}
