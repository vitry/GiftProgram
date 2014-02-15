using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace myLottery
{
    public partial class Form2 : Form
    {
        public String jpIcon;
        public Form2(String str, int parentFormWidth)
        {
            InitializeComponent();
            jpIcon = str;
            String path = System.Windows.Forms.Application.StartupPath + "\\";
            this.BackgroundImage = Image.FromFile(path + jpIcon);
            this.BackgroundImageLayout = ImageLayout.Stretch;
            this.WindowState = FormWindowState.Normal;
            this.StartPosition = FormStartPosition.Manual;
            this.FormBorderStyle = FormBorderStyle.None;
            this.Width = 400;
            this.Height = 300;
            this.TopMost = true;
            this.Top = 50;
            this.Left = (parentFormWidth - 400)/2;
        }

        private void Form1_DoubleClick(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
        }
    }
}
