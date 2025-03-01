using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using 插件.Properties;

namespace 插件.MyForm
{
    public partial class 聚光灯设置 : Form
    {
        public 聚光灯设置()
        {
            InitializeComponent();
        }

        private Color _Color = Settings.Default.聚光灯颜色;

        private void RadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            Settings.Default.聚光灯状态 = "1";
            PictureBox1.BackColor = this.BackColor;
            PictureBox2.BackColor = _Color;
        }

        private void RadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            Settings.Default.聚光灯状态 = "2";
            PictureBox2.BackColor = this.BackColor;
            PictureBox1.BackColor = _Color;
        }

        private void RadioButton3_CheckedChanged(object sender, EventArgs e)
        {
            Settings.Default.聚光灯状态 = "3";
            PictureBox1.BackColor = _Color;
            PictureBox2.BackColor = _Color;
        }

        private void TrackBar1_Scroll(object sender, EventArgs e)
        {
            Settings.Default.透明度 = TrackBar1.Value;
        }

        private void Label1_Click(object sender, EventArgs e)
        {
            try
            {
                ColorDialog colorDialog = new ColorDialog();

                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    Settings.Default.聚光灯颜色 = colorDialog.Color;
                    _Color = colorDialog.Color;
                    Settings.Default.Save();
                }
            }
            catch
            {
            }
        }
    }
}