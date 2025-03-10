namespace 插件.MyForm
{
    partial class 数据对比
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.colorComboBox = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.导出不同项 = new System.Windows.Forms.Button();
            this.导出相同项 = new System.Windows.Forms.Button();
            this.清除标识 = new System.Windows.Forms.Button();
            this.不同项 = new System.Windows.Forms.Button();
            this.相同项 = new System.Windows.Forms.Button();
            this.相同项Text = new System.Windows.Forms.TextBox();
            this.区域二Text = new System.Windows.Forms.TextBox();
            this.区域一Text = new System.Windows.Forms.TextBox();
            this.退出 = new System.Windows.Forms.Button();
            this.对比数据 = new System.Windows.Forms.Button();
            this.pictureBox2 = new System.Windows.Forms.PictureBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.区域2Box = new System.Windows.Forms.TextBox();
            this.区域1Box = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.colorComboBox);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.导出不同项);
            this.groupBox1.Controls.Add(this.导出相同项);
            this.groupBox1.Controls.Add(this.清除标识);
            this.groupBox1.Controls.Add(this.不同项);
            this.groupBox1.Controls.Add(this.相同项);
            this.groupBox1.Controls.Add(this.相同项Text);
            this.groupBox1.Controls.Add(this.区域二Text);
            this.groupBox1.Controls.Add(this.区域一Text);
            this.groupBox1.Controls.Add(this.退出);
            this.groupBox1.Controls.Add(this.对比数据);
            this.groupBox1.Controls.Add(this.pictureBox2);
            this.groupBox1.Controls.Add(this.pictureBox1);
            this.groupBox1.Controls.Add(this.区域2Box);
            this.groupBox1.Controls.Add(this.区域1Box);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Font = new System.Drawing.Font("宋体", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox1.Location = new System.Drawing.Point(2, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(613, 541);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "数据区域";
            // 
            // colorComboBox
            // 
            this.colorComboBox.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.colorComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.colorComboBox.FormattingEnabled = true;
            this.colorComboBox.Location = new System.Drawing.Point(483, 66);
            this.colorComboBox.Name = "colorComboBox";
            this.colorComboBox.Size = new System.Drawing.Size(97, 21);
            this.colorComboBox.TabIndex = 19;
            // 
            // label5
            // 
            this.label5.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.Location = new System.Drawing.Point(419, 107);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(59, 23);
            this.label5.TabIndex = 18;
            this.label5.Text = "相同项";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label4
            // 
            this.label4.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(211, 107);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(84, 23);
            this.label4.TabIndex = 17;
            this.label4.Text = "区域二独有";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(15, 107);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(84, 23);
            this.label3.TabIndex = 16;
            this.label3.Text = "区域一独有";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // 导出不同项
            // 
            this.导出不同项.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.导出不同项.Location = new System.Drawing.Point(503, 497);
            this.导出不同项.Name = "导出不同项";
            this.导出不同项.Size = new System.Drawing.Size(95, 23);
            this.导出不同项.TabIndex = 15;
            this.导出不同项.Text = "导出不同项&F";
            this.导出不同项.UseVisualStyleBackColor = true;
            this.导出不同项.Click += new System.EventHandler(this.导出不同项_Click);
            // 
            // 导出相同项
            // 
            this.导出相同项.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.导出相同项.Location = new System.Drawing.Point(379, 497);
            this.导出相同项.Name = "导出相同项";
            this.导出相同项.Size = new System.Drawing.Size(95, 23);
            this.导出相同项.TabIndex = 14;
            this.导出相同项.Text = "导出相同项&E";
            this.导出相同项.UseVisualStyleBackColor = true;
            this.导出相同项.Click += new System.EventHandler(this.导出相同项_Click);
            // 
            // 清除标识
            // 
            this.清除标识.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.清除标识.Location = new System.Drawing.Point(255, 497);
            this.清除标识.Name = "清除标识";
            this.清除标识.Size = new System.Drawing.Size(95, 23);
            this.清除标识.TabIndex = 13;
            this.清除标识.Text = "清除标识&C";
            this.清除标识.UseVisualStyleBackColor = true;
            this.清除标识.Click += new System.EventHandler(this.清除标识_Click);
            // 
            // 不同项
            // 
            this.不同项.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.不同项.Location = new System.Drawing.Point(131, 497);
            this.不同项.Name = "不同项";
            this.不同项.Size = new System.Drawing.Size(95, 23);
            this.不同项.TabIndex = 12;
            this.不同项.Text = "标识不同项&D";
            this.不同项.UseVisualStyleBackColor = true;
            this.不同项.Click += new System.EventHandler(this.不同项_Click);
            // 
            // 相同项
            // 
            this.相同项.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.相同项.Location = new System.Drawing.Point(7, 497);
            this.相同项.Name = "相同项";
            this.相同项.Size = new System.Drawing.Size(95, 23);
            this.相同项.TabIndex = 11;
            this.相同项.Text = "标识相同项&S";
            this.相同项.UseVisualStyleBackColor = true;
            this.相同项.Click += new System.EventHandler(this.相同项_Click);
            // 
            // 相同项Text
            // 
            this.相同项Text.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.相同项Text.Location = new System.Drawing.Point(422, 133);
            this.相同项Text.Multiline = true;
            this.相同项Text.Name = "相同项Text";
            this.相同项Text.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.相同项Text.Size = new System.Drawing.Size(176, 348);
            this.相同项Text.TabIndex = 10;
            // 
            // 区域二Text
            // 
            this.区域二Text.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.区域二Text.Location = new System.Drawing.Point(214, 133);
            this.区域二Text.Multiline = true;
            this.区域二Text.Name = "区域二Text";
            this.区域二Text.Size = new System.Drawing.Size(176, 348);
            this.区域二Text.TabIndex = 9;
            // 
            // 区域一Text
            // 
            this.区域一Text.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.区域一Text.Location = new System.Drawing.Point(10, 133);
            this.区域一Text.Multiline = true;
            this.区域一Text.Name = "区域一Text";
            this.区域一Text.Size = new System.Drawing.Size(176, 348);
            this.区域一Text.TabIndex = 8;
            // 
            // 退出
            // 
            this.退出.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.退出.Location = new System.Drawing.Point(397, 62);
            this.退出.Name = "退出";
            this.退出.Size = new System.Drawing.Size(77, 23);
            this.退出.TabIndex = 7;
            this.退出.Text = "退出";
            this.退出.UseVisualStyleBackColor = true;
            this.退出.Click += new System.EventHandler(this.退出_Click);
            // 
            // 对比数据
            // 
            this.对比数据.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.对比数据.Location = new System.Drawing.Point(302, 62);
            this.对比数据.Name = "对比数据";
            this.对比数据.Size = new System.Drawing.Size(71, 23);
            this.对比数据.TabIndex = 6;
            this.对比数据.Text = "对比数据";
            this.对比数据.UseVisualStyleBackColor = true;
            this.对比数据.Click += new System.EventHandler(this.对比数据_Click);
            // 
            // pictureBox2
            // 
            this.pictureBox2.Image = global::插件.Properties.Resources.选择;
            this.pictureBox2.Location = new System.Drawing.Point(227, 70);
            this.pictureBox2.Name = "pictureBox2";
            this.pictureBox2.Size = new System.Drawing.Size(28, 26);
            this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox2.TabIndex = 5;
            this.pictureBox2.TabStop = false;
            this.pictureBox2.Click += new System.EventHandler(this.pictureBox2_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::插件.Properties.Resources.选择;
            this.pictureBox1.Location = new System.Drawing.Point(228, 35);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(28, 26);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 4;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // 区域2Box
            // 
            this.区域2Box.Font = new System.Drawing.Font("宋体", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.区域2Box.Location = new System.Drawing.Point(88, 71);
            this.区域2Box.Name = "区域2Box";
            this.区域2Box.Size = new System.Drawing.Size(139, 25);
            this.区域2Box.TabIndex = 3;
            // 
            // 区域1Box
            // 
            this.区域1Box.Font = new System.Drawing.Font("宋体", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.区域1Box.Location = new System.Drawing.Point(89, 35);
            this.区域1Box.Name = "区域1Box";
            this.区域1Box.Size = new System.Drawing.Size(139, 25);
            this.区域1Box.TabIndex = 2;
            // 
            // label2
            // 
            this.label2.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(6, 74);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(76, 23);
            this.label2.TabIndex = 1;
            this.label2.Text = "区域二(T)";
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(6, 38);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(77, 23);
            this.label1.TabIndex = 0;
            this.label1.Text = "区域一(O)";
            // 
            // 数据对比
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(618, 544);
            this.Controls.Add(this.groupBox1);
            this.MaximumSize = new System.Drawing.Size(634, 583);
            this.MinimumSize = new System.Drawing.Size(634, 583);
            this.Name = "数据对比";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "数据对比";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.数据对比_FormClosed);
            this.Load += new System.EventHandler(this.Form2_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button 退出;
        private System.Windows.Forms.Button 对比数据;
        private System.Windows.Forms.PictureBox pictureBox2;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.TextBox 区域2Box;
        private System.Windows.Forms.TextBox 区域1Box;
        private System.Windows.Forms.Button 导出不同项;
        private System.Windows.Forms.Button 导出相同项;
        private System.Windows.Forms.Button 清除标识;
        private System.Windows.Forms.Button 不同项;
        private System.Windows.Forms.Button 相同项;
        private System.Windows.Forms.TextBox 相同项Text;
        private System.Windows.Forms.TextBox 区域二Text;
        private System.Windows.Forms.TextBox 区域一Text;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox colorComboBox;
    }
}