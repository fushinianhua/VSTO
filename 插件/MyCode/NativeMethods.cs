using System.Runtime.InteropServices;
using System;

public class NativeMethods
{
    // 获取当前焦点窗口
    [DllImport("user32.dll")]
    public static extern IntPtr GetForegroundWindow();

    // 获取窗口的进程 ID
    [DllImport("user32.dll", SetLastError = true)]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    // 设置全局钩子
    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

    // 卸载钩子
    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool UnhookWindowsHookEx(IntPtr hhk);

    // 调用下一个钩子
    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

    // 获取当前模块句柄
    [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    public static extern IntPtr GetModuleHandle(string lpModuleName);

    // 钩子类型：低级鼠标钩子
    public const int WH_MOUSE_LL = 14;

    // 鼠标滚轮消息
    public const int WM_MOUSEWHEEL = 0x020A;

    // 鼠标中键按下消息
    public const int WM_MBUTTONDOWN = 0x0207;

    // 委托：低级鼠标钩子回调函数
    public delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
    // 设置事件钩子
    [DllImport("user32.dll")]
    public static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess, uint idThread, uint dwFlags);

    // 卸载事件钩子
    [DllImport("user32.dll")]
    public static extern bool UnhookWinEvent(IntPtr hWinEventHook);

    // 委托：事件回调函数
    public delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

    // 事件常量：窗口焦点切换事件
    public const uint EVENT_SYSTEM_FOREGROUND = 0x0003;

    // 钩子标志：监听所有进程和线程
    public const uint WINEVENT_OUTOFCONTEXT = 0x0000;
}