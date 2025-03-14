using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace 插件
{
    internal class 窗口显示API
    {
        /// <summary>
        /// 设置窗口的显示状态
        /// </summary>
        /// <param name="hWnd">窗口句柄（窗口的唯一标识符）</param>
        /// <param name="nCmdShow">
        /// 窗口显示命令：
        ///   SW_HIDE = 0            （隐藏窗口）
        ///   SW_SHOWNORMAL = 1      （正常显示并激活窗口）
        ///   SW_SHOWMINIMIZED = 2   （最小化窗口）
        ///   SW_SHOWMAXIMIZED = 3    （最大化窗口）
        ///   SW_RESTORE = 9         （恢复窗口到正常大小和位置）
        /// </param>
        /// <returns>
        /// 操作结果：
        ///   true  - 窗口之前可见
        ///   false - 窗口之前被隐藏
        /// </returns>
        /// <remarks>
        /// 示例：WindowsAPI.ShowWindow(hWnd, WindowsAPI.SW_SHOWNORMAL);
        /// </remarks>
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        /// <summary>
        /// 将指定窗口设置为前台窗口并激活
        /// </summary>
        /// <param name="hWnd">目标窗口句柄</param>
        /// <returns>
        /// 操作结果：
        ///   true  - 成功将窗口置顶
        ///   false - 操作失败（可能因窗口不可用或权限不足）
        /// </returns>
        /// <remarks>
        /// 注意：
        /// 1. 窗口必须处于可见状态才能置顶
        /// 2. 进程必须具有前台权限（用户当前交互的进程）
        /// </remarks>
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        #region 窗口状态常量
        /// <summary>
        /// 隐藏窗口并激活其他窗口
        /// </summary>
        public const int SW_HIDE = 0;

        /// <summary>
        /// 正常显示并激活窗口（默认状态）
        /// </summary>
        public const int SW_SHOWNORMAL = 1;

        /// <summary>
        /// 最小化窗口并激活下一个顶层窗口
        /// </summary>
        public const int SW_SHOWMINIMIZED = 2;

        /// <summary>
        /// 最大化指定窗口
        /// </summary>
        public const int SW_SHOWMAXIMIZED = 3;

        /// <summary>
        /// 恢复窗口到正常大小和位置（与 SW_SHOWNORMAL 相同）
        /// </summary>
        public const int SW_RESTORE = 9;
        #endregion
    }
}
