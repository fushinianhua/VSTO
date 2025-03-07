using System.Runtime.InteropServices;
using System.Windows.Forms;
using System;

public class FocusMonitor
{
    private static IntPtr _hookID = IntPtr.Zero; // 钩子句柄
    private static NativeMethods.WinEventDelegate _delegate = WinEventCallback; // 回调函数

    // 启动监控
    public static void Start()
    {
        if (_hookID == IntPtr.Zero)
        {
            // 监听所有进程和线程的窗口焦点切换事件
            _hookID = NativeMethods.SetWinEventHook(
                NativeMethods.EVENT_SYSTEM_FOREGROUND, // 事件类型：窗口焦点切换
                NativeMethods.EVENT_SYSTEM_FOREGROUND, // 事件类型：窗口焦点切换
                IntPtr.Zero, // 监听所有模块
                _delegate, // 回调函数
                0, // 监听所有进程
                0, // 监听所有线程
                NativeMethods.WINEVENT_OUTOFCONTEXT // 标志：异步回调
            );

            if (_hookID == IntPtr.Zero)
            {
                MessageBox.Show("Failed to set window event hook!");
            }
            else
            {
                MessageBox.Show("Window event hook installed successfully!");
            }
        }
    }

    // 停止监控
    public static void Stop()
    {
        if (_hookID != IntPtr.Zero)
        {
            NativeMethods.UnhookWinEvent(_hookID);
            _hookID = IntPtr.Zero;
           // MessageBox.Show("Window event hook uninstalled!");
        }
    }

    // 回调函数：窗口焦点切换事件
    private static void WinEventCallback(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime)
    {
        if (eventType == NativeMethods.EVENT_SYSTEM_FOREGROUND)
        {
            // 获取窗口进程 ID
            uint processId;
            NativeMethods.GetWindowThreadProcessId(hwnd, out processId);

            // 获取进程名称
            var process = System.Diagnostics.Process.GetProcessById((int)processId);
            if (process.ProcessName.Equals("EXCEL", StringComparison.OrdinalIgnoreCase))
            {
                // 如果是 Excel 窗口，启动鼠标钩子
                //MessageBox.Show("Switched to Excel window, starting mouse hook...");
                MouseHook.Start();
            }
            else
            {
                // 如果不是 Excel 窗口，停止鼠标钩子
              //  MessageBox.Show("Switched to non-Excel window, stopping mouse hook...");
                MouseHook.Stop();
            }
        }
    }

    // 获取窗口进程 ID
    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
}