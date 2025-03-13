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
            this.数据 = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.数据)).BeginInit();
            this.SuspendLayout();
            // 
            // 数据
            // 
            this.数据.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.数据.Location = new System.Drawing.Point(12, 12);
            this.数据.Name = "数据";
            this.数据.RowTemplate.Height = 23;
            this.数据.Size = new System.Drawing.Size(970, 508);
            this.数据.TabIndex = 0;
            // 
            // 导出数据
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(994, 532);
            this.Controls.Add(this.数据);
            this.Name = "导出数据";
            this.Text = "导出数据";
            ((System.ComponentModel.ISupportInitialize)(this.数据)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView 数据;
    }
}