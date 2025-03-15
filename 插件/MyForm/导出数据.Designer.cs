namespace 插件.MyForm
{
    partial class 导出数据
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.CheckList = new System.Windows.Forms.CheckedListBox();
            this.文件导出 = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // CheckList
            // 
            this.CheckList.CheckOnClick = true;
            this.CheckList.FormattingEnabled = true;
            this.CheckList.Location = new System.Drawing.Point(8, 17);
            this.CheckList.Margin = new System.Windows.Forms.Padding(0);
            this.CheckList.MultiColumn = true;
            this.CheckList.Name = "CheckList";
            this.CheckList.Size = new System.Drawing.Size(543, 84);
            this.CheckList.TabIndex = 0;
            // 
            // 文件导出
            // 
            this.文件导出.Location = new System.Drawing.Point(473, 341);
            this.文件导出.Name = "文件导出";
            this.文件导出.Size = new System.Drawing.Size(63, 31);
            this.文件导出.TabIndex = 1;
            this.文件导出.Text = "导出";
            this.文件导出.UseVisualStyleBackColor = true;
            this.文件导出.Click += new System.EventHandler(this.文件导出_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.CheckList);
            this.groupBox1.Location = new System.Drawing.Point(3, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(558, 113);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "选择需要导出的列";
            // 
            // 导出数据
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(564, 384);
            this.Controls.Add(this.文件导出);
            this.Controls.Add(this.groupBox1);
            this.Name = "导出数据";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "导出数据";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.导出数据_FormClosed);
            this.Load += new System.EventHandler(this.导出数据_Load);
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.CheckedListBox CheckList;
        private System.Windows.Forms.Button 文件导出;
        private System.Windows.Forms.GroupBox groupBox1;
    }
}