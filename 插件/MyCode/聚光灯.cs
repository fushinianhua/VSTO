using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using 插件.Properties;
using Excel = Microsoft.Office.Interop.Excel;

namespace 插件.MyForm
{
    internal class 聚光灯
    {
        private readonly Excel.Application _excelApp;
        private Excel.Range _lastHighlightedRange;
        private Color _highlightColor = Settings.Default.聚光灯颜色;
        private const int VisibleRowLimit = 100;
        private const int VisibleColLimit = 50;

        public 聚光灯(Excel.Application excelApp)
        {
            _excelApp = excelApp;
            SubscribeEvents();
        }

        private void SubscribeEvents()
        {
            _excelApp.SheetSelectionChange += OnSelectionChanged;
        }

        public void UnsubscribeEvents()
        {
            _excelApp.SheetSelectionChange -= OnSelectionChanged;
            ClearHighlight();
        }

        public void SetHighlightColor(Color newColor)
        {
            _highlightColor = Color.FromArgb(80, newColor);
            UpdateHighlight();
        }

        private void OnSelectionChanged(object Sh, Excel.Range Target)
        {
            try
            {
                _excelApp.ScreenUpdating = false;
                ClearHighlight();
                ApplyHighlight(Target);
            }
            finally
            {
                _excelApp.ScreenUpdating = true;
            }
        }

        private void ClearHighlight()
        {
            if (_lastHighlightedRange == null) return;

            foreach (Excel.Range area in _lastHighlightedRange.Areas)
            {
                try
                {
                    area.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                }
                catch (COMException) { /* 处理合并单元格 */ }
                finally
                {
                    Marshal.ReleaseComObject(area);
                }
            }

            Marshal.ReleaseComObject(_lastHighlightedRange);
            _lastHighlightedRange = null;
        }

        private void ApplyHighlight(Excel.Range target)
        {
            var visibleRange = _excelApp.ActiveWindow.VisibleRange;

            // 计算行范围
            var rowsRange = CalculateRowRange(target, visibleRange);
            // 计算列范围
            var colsRange = CalculateColumnRange(target, visibleRange);

            _lastHighlightedRange = _excelApp.Union(rowsRange, colsRange);
            _lastHighlightedRange.Interior.Color = _highlightColor;

            Marshal.ReleaseComObject(rowsRange);
            Marshal.ReleaseComObject(colsRange);
            Marshal.ReleaseComObject(visibleRange);
        }

        private Excel.Range CalculateRowRange(Excel.Range target, Excel.Range visibleRange)
        {
            int firstRow = Math.Max(target.Row, visibleRange.Row);
            int lastRow = Math.Min(target.Row + target.Rows.Count - 1,
                visibleRange.Row + VisibleRowLimit);

            // 创建原始行范围
            Excel.Range fullRowRange = _excelApp.Range[
                _excelApp.Cells[firstRow, 1],
                _excelApp.Cells[lastRow, _excelApp.Columns.Count]
            ];

            // 使用Application.Intersect获取可见区域交集
            Excel.Range resultRange = _excelApp.Intersect(fullRowRange, visibleRange);

            // 释放中间对象
            Marshal.ReleaseComObject(fullRowRange);

            // 处理空范围情况
            if (resultRange == null)
            {
                return _excelApp.Range[_excelApp.Cells[1, 1], _excelApp.Cells[1, 1]]; // 返回虚拟范围
            }

            return resultRange;
        }

        private Excel.Range CalculateColumnRange(Excel.Range target, Excel.Range visibleRange)
        {
            int firstCol = Math.Max(target.Column, visibleRange.Column);
            int lastCol = Math.Min(target.Column + target.Columns.Count - 1,
                visibleRange.Column + VisibleRowLimit);
            Excel.Range fullRowRange = _excelApp.Range[
                _excelApp.Cells[1, firstCol],
                _excelApp.Cells[_excelApp.Rows.Count, lastCol]
            ];
            Excel.Range resultRange = _excelApp.Intersect(fullRowRange, visibleRange);

            // 释放中间对象
            Marshal.ReleaseComObject(fullRowRange);

            // 处理空范围情况
            if (resultRange == null)
            {
                return _excelApp.Range[_excelApp.Cells[1, 1], _excelApp.Cells[1, 1]]; // 返回虚拟范围
            }

            return resultRange;
        }

        private void UpdateHighlight()
        {
            if (_excelApp.Selection is Excel.Range currentRange)
            {
                ClearHighlight();
                ApplyHighlight(currentRange);
            }
        }
    }
}