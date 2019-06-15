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
    public partial class Frm_Login : Form
    {
        public Frm_Login()
        {
            InitializeComponent();
            this.skinEngine1.SkinFile = "../../Skins/CalmnessColor1.ssk";
            this.skinEngine1.SkinAllForm = true;
        }

        private void Form1_Load(object sender, EventArgs e)
        {


        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string username = txt_username.Text.Trim();
            string pwd = txt_Pwd.Text.Trim();
            SDSService.SDSServiceSoapClient sds = new ProcessControl.SDSService.SDSServiceSoapClient();
            DataSet ds = sds.MobileLogin(username, pwd);
            if (ds.Tables[0].Rows.Count > 0)
            {
                //MessageBox.Show("登录成功!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                LoginInfo.usercode = ds.Tables[0].Rows[0]["code"].ToString();
                LoginInfo.username = ds.Tables[0].Rows[0]["name"].ToString();
                LoginInfo.depotcode = ds.Tables[0].Rows[0]["depot"].ToString();
                LoginInfo.shop = ds.Tables[0].Rows[0]["shop"].ToString();
                Frm_Main fm = new Frm_Main();
                fm.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("账号或密码错误!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }

        }
    }
}
