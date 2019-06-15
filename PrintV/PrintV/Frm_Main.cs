using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ProcessControl
{
    public partial class Frm_Main : Form
    {
        public Frm_Main()
        {
            InitializeComponent();
            this.skinEngine1.SkinFile = "../../Skins/CalmnessColor1.ssk";
            this.skinEngine1.SkinAllForm = true;
        }

        private void toolStripLabel16_Click(object sender, EventArgs e)
        {
            DialogResult button = MessageBox.Show("确定要退出系统?", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk);

            if (button == DialogResult.Yes)
            {
                Application.ExitThread();
            }
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            //FrmLine fm2 = new FrmLine();
            //fm2.ShowDialog();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            tool_NowTime.Text = DateTime.Now.ToString();
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            //tool_UserName.Text = LoginInfo.username;
        }
    }
}
