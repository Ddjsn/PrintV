using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace ProcessControl
{
    public partial class FrmLine : Form
    {
        SDSService.SDSServiceSoapClient sds = new ProcessControl.SDSService.SDSServiceSoapClient();
        public FrmLine()
        {
            InitializeComponent();
            this.skinEngine1.SkinFile = "../../Skins/CalmnessColor1.ssk";
            this.skinEngine1.SkinAllForm = true;
        }

        private void toolStripLabel16_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {

        }



        private void textBox2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                //1.读取销售单
                KeyGetSell();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //1.读取销售单
            KeyGetSell();
        }

        private void KeyGetSell()
        {
            if (textBox2.Text.Trim() == "") {
                MessageBox.Show("请输入销售单号!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            else
            {
                DataSet ds01 = sds.GetSellMessage(textBox2.Text.Trim());
                if (ds01.Tables[0].Rows.Count > 0)
                {
                    textBox4.Text = ds01.Tables[0].Rows[0]["Code"].ToString();
                    textBox5.Text = ds01.Tables[0].Rows[0]["vname"].ToString();
                    textBox6.Text = ds01.Tables[0].Rows[0]["vDate"].ToString();
                    textBox7.Text = ds01.Tables[0].Rows[0]["vaname"].ToString();
                    textBox8.Text = ds01.Tables[0].Rows[0]["cname"].ToString();
                    textBox9.Text = ds01.Tables[0].Rows[0]["cdate"].ToString();
                    textBox10.Text = ds01.Tables[0].Rows[0]["Expname"].ToString();
                    textBox11.Text = ds01.Tables[0].Rows[0]["rmb"].ToString();
                    textBox12.Text = ds01.Tables[0].Rows[0]["scname"].ToString();
                    textBox13.Text = ds01.Tables[0].Rows[0]["state"].ToString();
                    textBox14.Text = ds01.Tables[0].Rows[0]["goodsType"].ToString();
                    textBox15.Text = ds01.Tables[0].Rows[0]["goodsType2"].ToString();

                    txtscode.Text = ds01.Tables[0].Rows[0]["scode"].ToString();

                    label4.Text = "1.已确定销售单号：" + ds01.Tables[0].Rows[0]["Code"].ToString() + ";请开始录入序列号！";
                    textBox3.Focus();
                }
                else
                {
                    MessageBox.Show("未找到销售单号!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
            }
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            ClearData();
        }

        private void textBox3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                LineOver();
            }
        }

        private void LineOver()
        {
            if (textBox4.Text != "")
            {
                if (textBox3.Text != "")
                {
                    try
                    {
                        //1.检测SN
                        string checkMsg = sds.CheckSN(textBox3.Text.Trim(), textBox4.Text.Trim(), txtscode.Text.Trim());
                        if (checkMsg.Contains("成功"))
                        {
                            label5.Text = checkMsg;
                            //检测是否还需要录入SN
                            DataSet Ds01 = sds.GetSNNumber(textBox4.Text.Trim());
                            //开始转打包和出库
                            if (Ds01.Tables[0].Rows[0][0].ToString().Trim() == "0")
                            {
                                label6.Text = "序列号扫描完成，开始自动生成出入库。";
                                DataSet Ds02 = sds.AutoSendGoods(textBox4.Text);
                                label7.Text = "自动生成出入库:" + Ds02.Tables[0].Rows[0][0].ToString() + ";损益配件数:" + Ds02.Tables[1].Rows[0]["cc"].ToString() + "";
                                //开始转打包
                                string SellFine = sds.AutoSellFiner(textBox4.Text, LoginInfo.username);
                                if (SellFine.Contains("成功"))
                                {
                                    label8.Text = SellFine;
                                    //开始出库
                                    string senMsg = sds.AutoSellSend(textBox4.Text, txtscode.Text, LoginInfo.username);
                                    if (senMsg.Contains("成功"))
                                    {
                                        label9.Text = senMsg;
                                        label10.Text = textBox4.Text + "自动化完成;正在开始下一个订单.....";
                                        Thread.Sleep(1000);
                                        ClearData();
                                    }
                                    else
                                    {
                                        label9.Text = senMsg + ";该订单未完成出库，请返回ERP手动操作";
                                        DialogResult res = MessageBox.Show(SellFine + ";是否开始新的订单!", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk);
                                        if (res == DialogResult.Yes)
                                        {
                                            //清空数据
                                            ClearData();
                                        }
                                    }
                                }
                                else
                                {
                                    label8.Text = SellFine + ";该订单未完成转打包请返回ERP手动操作";
                                    DialogResult res = MessageBox.Show(SellFine + ";是否开始新的订单!", "提示", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk);
                                    if (res == DialogResult.Yes)
                                    {
                                        //清空数据
                                        ClearData();
                                    }
                                }
                            }
                            else
                            {
                                textBox3.Focus();
                            }
                        }
                        else
                        {
                            label5.Text = checkMsg;
                            MessageBox.Show("该订单扫入SN码异常请单独处理!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        }
                    }
                    catch (Exception ex) {
                        ClearData();
                        MessageBox.Show("(已自动清理数据开始新的单号!)程序发生了无法处理的异常,请告知管理员!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    }
                }
                else
                {
                    MessageBox.Show("序列号扫入异常!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                
            }
            else
            {
                MessageBox.Show("未找到销售单!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
        }
        //F1开始下一轮
        private void FrmLine_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F1)
            {
                ClearData();
            }
        }
        private void ClearData() {
            textBox2.Text = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";
            textBox13.Text = "";
            textBox14.Text = "";
            textBox15.Text = "";
            txtscode.Text = "";

            label4.Text = "";
            label5.Text = "";
            label6.Text = "";
            label7.Text = "";
            label8.Text = "";
            label9.Text = "";
            label10.Text = "";
            textBox2.Focus();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            LineOver();
        }
      
    }
}
