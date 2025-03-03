using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using 插件.Properties;

namespace 插件.MyForm
{
    internal class StaticClass
    {
        public static Application ExcelApp;
        public static string _聚光灯选择状态;
        public static bool _聚光开关状态;
        // 定义一个事件，当聚光灯颜色改变时触发
        public event EventHandler<ColorChangedEventArgs> SpotlightColorChanged;
        public event EventHandler<状态ChangedEventArgs> Spotlight状态Changed;
        public event EventHandler<开关状态ChangedEventArgs> 开关状态Changed;
        // 私有字段，用于存储聚光灯颜色
        private static Color _聚光灯颜色;

        // 公共属性，用于获取和设置聚光灯颜色
        public static Color 聚光灯颜色
        {
            get
            {
                return _聚光灯颜色;
            }
            set
            {
                if (_聚光灯颜色 != value)
                {
                    // 保存旧颜色
                    Color oldColor = _聚光灯颜色;
                    // 更新颜色
                    _聚光灯颜色 = value;
                    // 触发事件
                    OnSpotlightColorChanged(oldColor, value);
                }
            }
        }
        public static string 聚光灯状态
        {
            get {   return _聚光灯选择状态; }
            set { 
                if (_聚光灯选择状态 != value)
                {
                    string old状态 = _聚光灯选择状态;
                    _聚光灯选择状态 = value;
                    OnSpotlight状态Changed(old状态,value);
                } 
            }

        }
        public static bool 聚光开关状态
        {
            get { return _聚光开关状态; }
            set
            {
                if (_聚光开关状态 != value)
                {
                    bool old状态 = _聚光开关状态;
                    _聚光开关状态 = value;
                    On开关状态Changed(old状态, value);
                }
            }

        }

        // 静态构造函数，初始化聚光灯颜色
        static StaticClass()
        {
            聚光灯颜色 = Settings.Default.聚光灯颜色;
            聚光灯状态 = Settings.Default.聚光灯选择状态;
            聚光开关状态 = Settings.Default.聚光灯开关状态;

        }

        // 触发事件的方法
        protected static void OnSpotlightColorChanged(Color oldColor, Color newColor)
        {
            // 获取事件的订阅者
            EventHandler<ColorChangedEventArgs> handler = Instance.SpotlightColorChanged;
            if (handler != null)
            {
                // 创建事件参数对象
                ColorChangedEventArgs args = new ColorChangedEventArgs(oldColor, newColor);
                // 触发事件
                handler(Instance, args);
            }
        }
        protected static void OnSpotlight状态Changed(string oldColor, string newColor)
        {
            // 获取事件的订阅者
            EventHandler<状态ChangedEventArgs> handler = Instance.Spotlight状态Changed;
            if (handler != null)
            {
                // 创建事件参数对象
                状态ChangedEventArgs args = new 状态ChangedEventArgs(oldColor, newColor);
                // 触发事件
                handler(Instance, args);
            }
        }
        protected static void On开关状态Changed(bool oldColor, bool newColor)
        {
            // 获取事件的订阅者
            EventHandler<开关状态ChangedEventArgs> handler = Instance.开关状态Changed;
            if (handler != null)
            {
                // 创建事件参数对象
                开关状态ChangedEventArgs args = new 开关状态ChangedEventArgs(oldColor, newColor);
                // 触发事件
                handler(Instance, args);
            }
        }
        // 单例模式，确保只有一个 SpotlightColorManager 实例
        private static StaticClass _instance;
        public static StaticClass Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new StaticClass();
                }
                return _instance;
            }
        }
        // 定义事件参数类，用于传递旧颜色和新颜色
        public class ColorChangedEventArgs : EventArgs
        {
            public Color OldColor { get; private set; }
            public Color NewColor { get; private set; }

            public ColorChangedEventArgs(Color oldColor, Color newColor)
            {
                OldColor = oldColor;
                NewColor = newColor;
            }
        }
        public class 状态ChangedEventArgs : EventArgs
        {
            public string Old状态 { get; private set; }
            public string New状态 { get; private set; }

            public 状态ChangedEventArgs(string old状态, string new状态)
            {
                Old状态 = old状态;
                New状态 = new状态;
            }
        }
        public class 开关状态ChangedEventArgs : EventArgs
        {
            public bool Old状态 { get; private set; }
            public bool New状态 { get; private set; }

            public 开关状态ChangedEventArgs(bool old状态, bool new状态)
            {
                Old状态 = old状态;
                New状态 = new状态;
            }
        }
    }
}