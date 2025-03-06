using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace 插件.MyForm
{
    internal class 聚光灯
    {
        // 存储 Excel 应用程序对象的引用，在构造函数中初始化且不可更改
        private readonly Excel.Application _excelApp;
        // 存储上一次高亮显示的单元格范围
        private Excel.Range _lastHighlightedRange;

        // 定义可见行的最大限制数量
        private const int VisibleRowLimit = 100;
        // 定义可见列的最大限制数量
        private const int VisibleColLimit = 50;
        private bool _isSpotlightEnabled;

        /// <summary>
        /// 构造函数，用于初始化聚光灯功能类
        /// </summary>
        /// <param name="excelApp">Excel 应用程序对象，不能为 null</param>
        /// <exception cref="ArgumentNullException">当传入的 excelApp 为 null 时抛出该异常</exception>
        public 聚光灯(Excel.Application excelApp)
        {
            // 对传入的 excelApp 进行空值检查，若为 null 则抛出异常
            _excelApp = excelApp ?? throw new ArgumentNullException(nameof(excelApp), "Excel应用程序对象不能为null");
            _isSpotlightEnabled = StaticClass._聚光开关状态;
            // 订阅 Excel 的相关事件，以便在事件触发时执行相应操作
            SubscribeEvents();
        }

        /// <summary>
        /// 订阅 Excel 的工作表选择更改和窗口大小调整事件
        /// </summary>
        private void SubscribeEvents()
        {
            // 当工作表中的选择区域发生改变时，触发 OnSelectionChanged 方法
            _excelApp.SheetSelectionChange += OnSelectionChanged;
            // 当 Excel 窗口大小调整时，触发 ExcelApp_WindowResize 方法
            _excelApp.WindowResize += ExcelApp_WindowResize;
            //
         
            //
            StaticClass.Instance.SpotlightColorChanged += ColorValueChanged;
            //
            StaticClass.Instance.Spotlight状态Changed += 状态ValueChanged;
            StaticClass.Instance.开关状态Changed += 开关状态Changed;
            
        }

        private void 开关状态Changed(object sender, StaticClass.开关状态ChangedEventArgs e)
        {
            _isSpotlightEnabled = StaticClass.聚光开关状态;
            if (_isSpotlightEnabled)
            {

                // 当工作表中的选择区域发生改变时，触发 OnSelectionChanged 方法
                _excelApp.SheetSelectionChange += OnSelectionChanged;
                // 当 Excel 窗口大小调整时，触发 ExcelApp_WindowResize 方法
                _excelApp.WindowResize += ExcelApp_WindowResize;
                ApplyHighlight(_excelApp.Selection as Excel.Range);
            }
            else
            {
                // 当工作表中的选择区域发生改变时，触发 OnSelectionChanged 方法
                _excelApp.SheetSelectionChange -= OnSelectionChanged;
                // 当 Excel 窗口大小调整时，触发 ExcelApp_WindowResize 方法
                _excelApp.WindowResize -= ExcelApp_WindowResize;
                ClearHighlight();
            }

            }

        private void 状态ValueChanged(object sender, StaticClass.状态ChangedEventArgs e)
        {
            if (_isSpotlightEnabled)
            {
                DisableExcelUpdates();
                try
                {
                    // 清除上一次高亮显示的单元格颜色
                    ClearHighlight();
                    ApplyHighlight(_excelApp.Selection as Excel.Range);
                }
                finally
                {
                    // 恢复 Excel 的屏幕更新和事件触发
                    EnableExcelUpdates();
                }
            }
        }

        private void ColorValueChanged(object sender, StaticClass.ColorChangedEventArgs e)
        {
            if (_isSpotlightEnabled)
            {
                DisableExcelUpdates();
                try
                {
                    // 清除上一次高亮显示的单元格颜色
                    ClearHighlight();
                    ApplyHighlight(_excelApp.Selection as Excel.Range);
                }
                finally
                {
                    // 恢复 Excel 的屏幕更新和事件触发
                    EnableExcelUpdates();
                }
            }
        }

        /// <summary>
        /// 当 Excel 窗口大小调整时执行的方法
        /// </summary>
        /// <param name="Wb">当前工作簿对象</param>
        /// <param name="Wn">当前窗口对象</param>
        private void ExcelApp_WindowResize(Excel.Workbook Wb, Excel.Window Wn)
        {
            if (_isSpotlightEnabled)
            {
                // 清除上一次高亮显示的单元格颜色
                ClearHighlight();
                // 对当前活动单元格应用高亮显示
                ApplyHighlight(_excelApp.Selection as Excel.Range);
            }
        }

        /// <summary>
        /// 取消订阅 Excel 事件，并清除高亮显示
        /// </summary>
        public void UnsubscribeEvents()
        {
            // 取消订阅工作表选择更改事件
            _excelApp.SheetSelectionChange -= OnSelectionChanged;
            // 清除上一次高亮显示的单元格颜色
            ClearHighlight();
        }

        /// <summary>
        /// 设置聚光灯的高亮颜色，并更新高亮显示
        /// </summary>
        /// <param name="newColor">新的高亮颜色</param>
        public void SetHighlightColor(Color newColor)
        {
            if (_isSpotlightEnabled)
            {
                // 此处代码被注释，原本功能可能是设置静态类中存储的聚光灯颜色
                // StaticClass.聚光灯颜色 = Color.FromArgb(80, newColor);
                // 更新高亮显示，以应用新的颜色
                UpdateHighlight();
            }
        }

        /// <summary>
        /// 当工作表中的选择区域发生改变时执行的方法
        /// </summary>
        /// <param name="Sh">工作表对象</param>
        /// <param name="Target">新选择的单元格范围</param>
        private void OnSelectionChanged(object Sh, Excel.Range Target)
        {
            if (_isSpotlightEnabled)
            {
                // 禁用 Excel 的屏幕更新和事件触发，避免操作过程中界面闪烁和不必要的事件响应
                DisableExcelUpdates();
                try
                {
                    // 清除上一次高亮显示的单元格颜色
                    ClearHighlight();
                    // 对新选择的单元格范围应用高亮显示
                    ApplyHighlight(Target);
                }
                finally
                {
                    // 恢复 Excel 的屏幕更新和事件触发
                    EnableExcelUpdates();
                }
            }
        }

        /// <summary>
        /// 清除上一次高亮显示的单元格颜色
        /// </summary>
        private void ClearHighlight()
        {
            // 如果上一次高亮显示的范围为空，则直接返回
            if (_lastHighlightedRange == null) return;

            // 遍历上一次高亮显示范围中的每个区域
            foreach (Excel.Range area in _lastHighlightedRange.Areas)
            {
                try
                {
                    // 将该区域内单元格的背景颜色设置为无
                    area.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                }
                catch (COMException)
                {
                    // 当处理合并单元格时可能会抛出 COM 异常，此处记录日志方便调试
                   Debug.WriteLine("处理合并单元格时发生COM异常");
                }
                finally
                {
                    // 释放该区域的 COM 对象资源，避免内存泄漏
                    ReleaseComObject(area);
                }
            }

            // 释放上一次高亮显示范围的 COM 对象资源
            ReleaseComObject(_lastHighlightedRange);
            // 将上一次高亮显示范围置为 null
            _lastHighlightedRange = null;
        }

        /// <summary>
        /// 对指定的单元格范围应用高亮显示
        /// </summary>
        /// <param name="target">要应用高亮显示的单元格范围</param>
        private void ApplyHighlight(Excel.Range target)
        {
            if (target == null || !_isSpotlightEnabled) return;

            // 判断是否选择了整行或整列
            bool isWholeRow = target.Columns.Count == _excelApp.Columns.Count;
            bool isWholeColumn = target.Rows.Count == _excelApp.Rows.Count;
            if (isWholeRow || isWholeColumn)
            {
                return;
            }

            // 禁用 Excel 的屏幕更新和事件触发，避免操作过程中界面闪烁和不必要的事件响应
            DisableExcelUpdates();
            try
            {
                // 获取当前 Excel 窗口的可见范围
                var visibleRange = _excelApp.ActiveWindow.VisibleRange;
                try
                {
                    // 根据静态类中存储的聚光灯状态，选择不同的高亮显示方式
                    switch (StaticClass.聚光灯状态)
                    {
                        case "1":
                            // 只高亮显示目标单元格所在的行
                            _lastHighlightedRange = CalculateRowRange(target, visibleRange);
                            break;
                        case "2":
                            // 只高亮显示目标单元格所在的列
                            _lastHighlightedRange = CalculateColumnRange(target, visibleRange);
                            break;
                        case "3":
                            // 同时高亮显示目标单元格所在的行和列
                            var rowsRange = CalculateRowRange(target, visibleRange);
                            var colsRange = CalculateColumnRange(target, visibleRange);
                            try
                            {
                                _lastHighlightedRange = _excelApp.Union(rowsRange, colsRange);
                            }
                            finally
                            {
                                ReleaseComObject(rowsRange);
                                ReleaseComObject(colsRange);
                            }
                            break;
                        default:
                            // 当聚光灯状态为未知值时，记录日志并返回，不进行高亮显示操作
                            System.Diagnostics.Debug.WriteLine($"未知的聚光灯状态: {StaticClass.聚光灯状态}");
                            return;
                    }

                    // 将最终确定的高亮显示范围的背景颜色设置为静态类中存储的聚光灯颜色
                    _lastHighlightedRange.Interior.Color = StaticClass.聚光灯颜色;
                    target.Interior.Color = Color.White;
                }
                finally
                {
                    // 释放可见范围的 COM 对象资源
                    ReleaseComObject(visibleRange);
                }
            }
            catch (Exception ex)
            {
                // 当应用高亮显示过程中发生异常时，记录异常信息方便调试
                System.Diagnostics.Debug.WriteLine($"应用高亮显示时发生异常: {ex.Message}");
            }
            finally
            {
                // 恢复 Excel 的屏幕更新和事件触发
                EnableExcelUpdates();
            }
        }

        /// <summary>
        /// 计算目标单元格所在行的高亮显示范围
        /// </summary>
        /// <param name="target">目标单元格范围</param>
        /// <param name="visibleRange">当前 Excel 窗口的可见范围</param>
        /// <returns>计算得到的行高亮显示范围</returns>
        private Excel.Range CalculateRowRange(Excel.Range target, Excel.Range visibleRange)
        {
            // 计算起始行，取目标行和可见范围起始行的较大值
            int firstRow = Math.Max(target.Row, visibleRange.Row);
            // 计算结束行，取目标行加上目标行数减1和可见范围起始行加上可见行限制的较小值
            int lastRow = Math.Min(target.Row + target.Rows.Count - 1,
                visibleRange.Row + VisibleRowLimit);

            // 创建原始行范围，从起始行第一列到结束行最后一列
            Excel.Range fullRowRange = _excelApp.Range[_excelApp.Cells[firstRow, 1], _excelApp.Cells[lastRow, _excelApp.Columns.Count]];

            // 使用Intersect方法获取原始行范围和可见范围的交集
            Excel.Range resultRange = _excelApp.Intersect(fullRowRange, visibleRange);

            // 释放原始行范围的COM对象资源
            Marshal.ReleaseComObject(fullRowRange);

            // 处理交集为空的情况，返回一个虚拟范围
            if (resultRange == null)
            {
                return _excelApp.Range[_excelApp.Cells[1, 1], _excelApp.Cells[1, 1]];
            }

            return resultRange;
        }

        /// <summary>
        /// 计算目标单元格所在列的高亮显示范围
        /// </summary>
        /// <param name="target">目标单元格范围</param>
        /// <param name="visibleRange">当前 Excel 窗口的可见范围</param>
        /// <returns>计算得到的列高亮显示范围</returns>
        private Excel.Range CalculateColumnRange(Excel.Range target, Excel.Range visibleRange)
        {
            // 计算起始列，取目标列和可见范围起始列的较大值
            int firstCol = Math.Max(target.Column, visibleRange.Column);
            // 计算结束列，取目标列加上目标列数减1和可见范围起始列加上可见行限制的较小值
            int lastCol = Math.Min(target.Column + target.Columns.Count - 1,
                visibleRange.Column + VisibleRowLimit);

            // 创建原始列范围，从第一行起始列到最后一行结束列
            Excel.Range fullRowRange = _excelApp.Range[_excelApp.Cells[1, firstCol], _excelApp.Cells[_excelApp.Rows.Count, lastCol]];

            // 使用Intersect方法获取原始列范围和可见范围的交集
            Excel.Range resultRange = _excelApp.Intersect(fullRowRange, visibleRange);

            // 释放原始列范围的COM对象资源
            Marshal.ReleaseComObject(fullRowRange);

            // 处理交集为空的情况，返回一个虚拟范围
            if (resultRange == null)
            {
                return _excelApp.Range[_excelApp.Cells[1, 1], _excelApp.Cells[1, 1]];
            }

            return resultRange;
        }

        /// <summary>
        /// 更新高亮显示，清除之前的高亮并对当前选择的单元格范围应用高亮
        /// </summary>
        private void UpdateHighlight()
        {
            if (_isSpotlightEnabled)
            {
                // 检查当前选择的对象是否为 Excel 单元格范围
                if (_excelApp.Selection is Excel.Range currentRange)
                {
                    // 清除上一次高亮显示的单元格颜色
                    ClearHighlight();
                    // 对当前选择的单元格范围应用高亮显示
                    ApplyHighlight(currentRange);
                }
            }
        }

        /// <summary>
        /// 禁用 Excel 的屏幕更新和事件触发
        /// </summary>
        private void DisableExcelUpdates()
        {
            _excelApp.ScreenUpdating = false;
            _excelApp.DisplayAlerts = false;
            _excelApp.EnableEvents = false;
        }

        /// <summary>
        /// 启用 Excel 的屏幕更新和事件触发
        /// </summary>
        private void EnableExcelUpdates()
        {
            _excelApp.ScreenUpdating = true;
            _excelApp.DisplayAlerts = true;
            _excelApp.EnableEvents = true;
        }

        /// <summary>
        /// 释放 COM 对象资源
        /// </summary>
        /// <param name="obj">要释放的 COM 对象</param>
        private void ReleaseComObject(object obj)
        {
            // 检查对象是否为 COM 对象且不为 null
            if (obj != null && Marshal.IsComObject(obj))
            {
                // 释放 COM 对象资源
                Marshal.ReleaseComObject(obj);
            }
        }

        /// <summary>
        /// 开启聚光灯功能
        /// </summary>
        public void EnableSpotlight()
        {
            _isSpotlightEnabled = true;
            UpdateHighlight();
        }

        /// <summary>
        /// 关闭聚光灯功能
        /// </summary>
        public void DisableSpotlight()
        {
            _isSpotlightEnabled = false;
            ClearHighlight();
        }

        /// <summary>
        /// 获取或设置聚光灯的开关状态
        /// </summary>
        public bool IsSpotlightEnabled
        {
            get => _isSpotlightEnabled;
            set
            {
                if (value)
                {
                    EnableSpotlight();
                }
                else
                {
                    DisableSpotlight();
                }
            }
        }
    }
}