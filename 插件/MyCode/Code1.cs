using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace 插件.MyForm
{
    internal class Code1
    {
        private static Application exapp = StaticClass.ExcelApp;

        public static string StrtoW(string SrtText)
        {
            try
            {
                // 查找字符串中 "." 的位置
                int index = SrtText.IndexOf('.');
                if (index == -1)
                {
                    throw new ArgumentException("输入的字符串中不包含 '.'。");
                }

                // 提取列号部分并转换为长整型
                string colNumStr = SrtText.Substring(0, index);
                long colNum = long.Parse(colNumStr);

                // 获取对应列的单元格地址并去除行号部分
                Range cell = exapp.Cells[1, colNum];
                string address = cell.Address[false, false];
                string colLetter = address.Replace("1", "");

                // 释放 COM 对象
                System.Runtime.InteropServices.Marshal.ReleaseComObject(cell);

                return colLetter;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"发生错误: {ex.Message}");
                return null;
            }
        }
    }
}