using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
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

        private Color _Color;
        private string _状态;
        public Color Color
        {
            get { return _Color; }
            set {
                if (_Color != value)
                {
                    _Color = value;
                    ColorChanged(value);
                }
               
            }
        }
        private void ColorChanged(Color color)
        {
            if (RadioButton1.Checked)
            {
                PictureBox2.BackColor = color;
                PictureBox1.BackColor = this.BackColor;
            }
            if (RadioButton2.Checked)
            {
                PictureBox1.BackColor = color;
                PictureBox2.BackColor = this.BackColor;
            }

            if (RadioButton3.Checked)
            {
                PictureBox1.BackColor = color;
                PictureBox2.BackColor = color;
            }         
            Settings.Default.聚光灯颜色 = color;
           StaticClass.聚光灯颜色=color;
            Settings.Default.Save();
        }
        private void RadioButton1_CheckedChanged(object sender, EventArgs e)
        {
            _状态 = "1";
            PictureBox1.BackColor = this.BackColor;
            PictureBox2.BackColor = _Color;
        }

        private void RadioButton2_CheckedChanged(object sender, EventArgs e)
        {
            _状态 = "2";
            PictureBox2.BackColor = this.BackColor;
            PictureBox1.BackColor = _Color;
        }

        private void RadioButton3_CheckedChanged(object sender, EventArgs e)
        {
            _状态 = "3";
            PictureBox1.BackColor = _Color;
            PictureBox2.BackColor = _Color;
        }

        private void TrackBar1_Scroll(object sender, EventArgs e)
        {
            Settings.Default.透明度 = TrackBar1.Value;
            Settings.Default.Save();
        }

        private void Label1_Click(object sender, EventArgs e)
        {
            try
            {
                ColorDialog colorDialog = new ColorDialog();

                if (colorDialog.ShowDialog() == DialogResult.OK)
                {
                    Color = colorDialog.Color;                 
                }
            }
            catch
            {
            }
        }

        private void 聚光灯设置_Load(object sender, EventArgs e)
        {
            _Color = StaticClass.聚光灯颜色=Settings.Default.聚光灯颜色;
            _状态 = StaticClass.聚光灯状态= Settings.Default.聚光灯选择状态;
            if (_状态 == "1")
            {
                RadioButton1.Checked = true;
            }
            else if (_状态 == "2")
            {
                RadioButton2.Checked = true;
            }
            else
            {
                RadioButton3.Checked = true;
            }

        }
       
        private void Button1_Click(object sender, EventArgs e)
        {
            StaticClass.聚光灯颜色 = _Color;
            StaticClass.聚光灯状态 = _状态;
            Settings.Default.聚光灯选择状态 = _状态;
            Settings.Default.聚光灯颜色 = _Color;
            Settings.Default.Save();
            this.Close();
        }

        private void 聚光灯设置_FormClosed(object sender, FormClosedEventArgs e)
        {
           AMG. 聚光灯form=null;
        }
    }
}