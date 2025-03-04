using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

public class MouseWheelHandler
{
    private const int WH_MOUSE_LL = 14;
    private const int WM_MOUSEWHEEL = 0x020A;
    private static IntPtr _hookID = IntPtr.Zero;
    private static LowLevelMouseProc _proc;
    private static Excel.Application _excelApp;

    public static void Initialize(Excel.Application excelApp)
    {
        _excelApp = excelApp;
        _proc = HookCallback;
        _hookID = SetHook(_proc);
    }

    private static IntPtr SetHook(LowLevelMouseProc proc)
    {
        using (var curProcess = System.Diagnostics.Process.GetCurrentProcess())
        using (var curModule = curProcess.MainModule)
        {
            return SetWindowsHookEx(WH_MOUSE_LL, proc,
                GetModuleHandle(curModule.ModuleName), 0);
        }
    }

    private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

    private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0 && wParam == (IntPtr)WM_MOUSEWHEEL)
        {
            MSLLHOOKSTRUCT hookStruct = Marshal.PtrToStructure<MSLLHOOKSTRUCT>(lParam);
            // 获取鼠标位置并转换为Excel坐标
            Excel.Window window = _excelApp.ActiveWindow;
            Excel.Range visibleRange = window.VisibleRange;

            // 计算滚动方向（delta右移16位获取符号）
            int delta = (hookStruct.mouseData >> 16) & 0xffff;
            int scrollDirection = delta > 0 ? 1 : -1;

            // 获取当前聚光灯位置并更新
            UpdateSpotlight(scrollDirection, visibleRange);
        }
        return CallNextHookEx(_hookID, nCode, wParam, lParam);
    }

    private static void UpdateSpotlight(int direction, Excel.Range visibleRange)
    {
        // 假设当前聚光灯位于activeCell
        Excel.Range activeCell = _excelApp.ActiveCell;
        int newRow = activeCell.Row + direction;

        // 确保新位置在可见范围内
        if (newRow >= visibleRange.Row &&
            newRow <= visibleRange.Row + visibleRange.Rows.Count)
        {
            // 移动聚光灯到新位置
            Excel.Range newCell = activeCell.Worksheet.Cells[newRow, activeCell.Column];
            //ApplySpotlight(newCell);
        }
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct POINT { public int x; public int y; }

    [StructLayout(LayoutKind.Sequential)]
    private struct MSLLHOOKSTRUCT
    {
        public POINT pt;
        public int mouseData;
        public int flags;
        public int time;
        public IntPtr dwExtraInfo;
    }

    #region Windows API
    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr SetWindowsHookEx(int idHook,
        LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool UnhookWindowsHookEx(IntPtr hhk);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode,
        IntPtr wParam, IntPtr lParam);

    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern IntPtr GetModuleHandle(string lpModuleName);
    #endregion
}