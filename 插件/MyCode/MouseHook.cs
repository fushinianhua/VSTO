using System.Windows.Forms;
using System;

public class MouseHook
{
    private static IntPtr _hookID = IntPtr.Zero; // 钩子句柄
    private static NativeMethods.LowLevelMouseProc _proc = HookCallback; // 回调函数

    // 启动钩子
    public static void Start()
    {
        if (_hookID == IntPtr.Zero)
        {
            _hookID = SetHook(_proc);
        }
    }

    // 停止钩子
    public static void Stop()
    {
        if (_hookID != IntPtr.Zero)
        {
            NativeMethods.UnhookWindowsHookEx(_hookID);
            _hookID = IntPtr.Zero;
        }
    }

    private static IntPtr SetHook(NativeMethods.LowLevelMouseProc proc)
    {
        using (var curProcess = System.Diagnostics.Process.GetCurrentProcess())
        using (var curModule = curProcess.MainModule)
        {
            IntPtr hookID = NativeMethods.SetWindowsHookEx(NativeMethods.WH_MOUSE_LL, proc, NativeMethods.GetModuleHandle(curModule.ModuleName), 0);
            if (hookID == IntPtr.Zero)
            {
               
            }
            else
            {
                
            }
            return hookID;
        }
    }

    private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
    {
        if (nCode >= 0)
        {
            Console.WriteLine($"wParam: {wParam}, lParam: {lParam}");

            if (wParam == (IntPtr)NativeMethods.WM_MOUSEWHEEL)
            {
                // 获取滚动的行数
                short delta = (short)((wParam.ToInt64() >> 16) & 0xFFFF);
                if (delta == 0)
                {
                   // MessageBox.Show("buzhengq");
                }
                int lines = delta / 120; // 每行通常为 120 单位

                // 输出调试信息
                Console.WriteLine($"Delta: {delta}, Lines: {lines}");

                // 触发自定义方法，并传递滚动的行数
                OnMouseWheelScrolled(lines);
            }
        }
        return NativeMethods.CallNextHookEx(_hookID, nCode, wParam, lParam);
    }

    // 自定义事件：鼠标滚轮滚动（带行数参数）
    public static event Action<int> MouseWheelScrolled;
    private static void OnMouseWheelScrolled(int lines)
    {
        MouseWheelScrolled?.Invoke(lines);
    }
}