namespace 插件
{
    partial class AMG : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public AMG()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button2 = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.聚光灯 = this.Factory.CreateRibbonSplitButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.button4 = this.Factory.CreateRibbonButton();
            this.导入数据 = this.Factory.CreateRibbonButton();
            this.导出数据 = this.Factory.CreateRibbonButton();
            this.group5 = this.Factory.CreateRibbonGroup();
            this.menu1 = this.Factory.CreateRibbonMenu();
            this.升序 = this.Factory.CreateRibbonButton();
            this.降序 = this.Factory.CreateRibbonButton();
            this.筛选 = this.Factory.CreateRibbonButton();
            this.清除筛选 = this.Factory.CreateRibbonButton();
            this.重新应用 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            this.group5.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Groups.Add(this.group5);
            this.tab1.Label = "快递操作";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.button1);
            this.group1.Label = "查询";
            this.group1.Name = "group1";
            // 
            // button1
            // 
            this.button1.Label = "数据匹配";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.button2);
            this.group2.Label = "对比";
            this.group2.Name = "group2";
            // 
            // button2
            // 
            this.button2.Label = "数据对比";
            this.button2.Name = "button2";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.聚光灯);
            this.group3.Label = "聚光灯";
            this.group3.Name = "group3";
            // 
            // 聚光灯
            // 
            this.聚光灯.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.聚光灯.Image = global::插件.Properties.Resources.聚光灯开;
            this.聚光灯.Items.Add(this.button3);
            this.聚光灯.Label = "聚光灯";
            this.聚光灯.Name = "聚光灯";
            this.聚光灯.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.聚光灯_Click);
            // 
            // button3
            // 
            this.button3.Label = "聚光灯设置";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.button4);
            this.group4.Items.Add(this.导入数据);
            this.group4.Items.Add(this.导出数据);
            this.group4.Label = "工作簿操作";
            this.group4.Name = "group4";
            // 
            // button4
            // 
            this.button4.Label = "拆分工作表";
            this.button4.Name = "button4";
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button4_Click);
            // 
            // 导入数据
            // 
            this.导入数据.Label = "导入";
            this.导入数据.Name = "导入数据";
            this.导入数据.OfficeImageId = "ActiveXButton";
            this.导入数据.ShowImage = true;
            this.导入数据.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.导入数据_Click);
            // 
            // 导出数据
            // 
            this.导出数据.Label = "导出";
            this.导出数据.Name = "导出数据";
            this.导出数据.OfficeImageId = "ActiveXButton";
            this.导出数据.ShowImage = true;
            this.导出数据.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.导出数据_Click);
            // 
            // group5
            // 
            this.group5.Items.Add(this.menu1);
            this.group5.Name = "group5";
            // 
            // menu1
            // 
            this.menu1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu1.ImageName = "SortFilterMenu";
            this.menu1.Items.Add(this.升序);
            this.menu1.Items.Add(this.降序);
            this.menu1.Items.Add(this.筛选);
            this.menu1.Items.Add(this.清除筛选);
            this.menu1.Items.Add(this.重新应用);
            this.menu1.Label = "筛选和排序";
            this.menu1.Name = "menu1";
            this.menu1.OfficeImageId = "SortFilterMenu";
            this.menu1.ShowImage = true;
            // 
            // 升序
            // 
            this.升序.ImageName = "SortUp";
            this.升序.Label = "升序";
            this.升序.Name = "升序";
            this.升序.OfficeImageId = "SortUp";
            this.升序.ShowImage = true;
            this.升序.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button5_Click);
            // 
            // 降序
            // 
            this.降序.ImageName = "SortDown";
            this.降序.Label = "降序";
            this.降序.Name = "降序";
            this.降序.OfficeImageId = "SortDown";
            this.降序.ShowImage = true;
            this.降序.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click);
            // 
            // 筛选
            // 
            this.筛选.ImageName = "Filter";
            this.筛选.Label = "筛选";
            this.筛选.Name = "筛选";
            this.筛选.OfficeImageId = "Filter";
            this.筛选.ShowImage = true;
            this.筛选.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.筛选_Click);
            // 
            // 清除筛选
            // 
            this.清除筛选.ImageName = "FilterClearAllFilters";
            this.清除筛选.Label = "清除筛选";
            this.清除筛选.Name = "清除筛选";
            this.清除筛选.OfficeImageId = "FilterClearAllFilters";
            this.清除筛选.ShowImage = true;
            this.清除筛选.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.清除筛选_Click);
            // 
            // 重新应用
            // 
            this.重新应用.ImageName = "FilterReapply";
            this.重新应用.Label = "重新应用";
            this.重新应用.Name = "重新应用";
            this.重新应用.OfficeImageId = "FilterReapply";
            this.重新应用.ShowImage = true;
            this.重新应用.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.重新应用_Click);
            // 
            // AMG
            // 
            this.Name = "AMG";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.AMG_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.group5.ResumeLayout(false);
            this.group5.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSplitButton 聚光灯;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group5;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 升序;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 降序;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 重新应用;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 筛选;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 清除筛选;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 导入数据;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton 导出数据;
    }

    partial class ThisRibbonCollection
    {
        internal AMG AMG
        {
            get { return this.GetRibbon<AMG>(); }
        }
    }
}
