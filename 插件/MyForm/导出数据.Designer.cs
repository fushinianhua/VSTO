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
            this.PathText = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.WBnameText = new System.Windows.Forms.TextBox();
            this.WSnameText = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
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
            // PathText
            // 
            this.PathText.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.PathText.Location = new System.Drawing.Point(89, 201);
            this.PathText.Multiline = true;
            this.PathText.Name = "PathText";
            this.PathText.Size = new System.Drawing.Size(323, 30);
            this.PathText.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(12, 201);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 30);
            this.label1.TabIndex = 4;
            this.label1.Text = "保存地址";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(12, 258);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(83, 30);
            this.label2.TabIndex = 5;
            this.label2.Text = "工作薄名称";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(242, 258);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(83, 30);
            this.label3.TabIndex = 6;
            this.label3.Text = "工作表名称";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // WBnameText
            // 
            this.WBnameText.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.WBnameText.Location = new System.Drawing.Point(89, 258);
            this.WBnameText.Multiline = true;
            this.WBnameText.Name = "WBnameText";
            this.WBnameText.Size = new System.Drawing.Size(81, 30);
            this.WBnameText.TabIndex = 7;
            // 
            // WSnameText
            // 
            this.WSnameText.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.WSnameText.Location = new System.Drawing.Point(331, 258);
            this.WSnameText.Multiline = true;
            this.WSnameText.Name = "WSnameText";
            this.WSnameText.Size = new System.Drawing.Size(81, 30);
            this.WSnameText.TabIndex = 8;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(416, 201);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(38, 30);
            this.button1.TabIndex = 9;
            this.button1.Text = "浏览";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // 导出数据
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(564, 384);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.WSnameText);
            this.Controls.Add(this.WBnameText);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.PathText);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.文件导出);
            this.Controls.Add(this.groupBox1);
            this.Name = "导出数据";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "导出数据";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.导出数据_FormClosed);
            this.Load += new System.EventHandler(this.导出数据_Load);
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckedListBox CheckList;
        private System.Windows.Forms.Button 文件导出;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox PathText;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox WBnameText;
        private System.Windows.Forms.TextBox WSnameText;
        private System.Windows.Forms.Button button1;
    }
}