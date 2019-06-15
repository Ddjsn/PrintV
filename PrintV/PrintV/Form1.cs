using Microsoft.Reporting.WinForms;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace PrintV
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private int m_currentPageIndex;
        private IList<Stream> m_streams;
        private void Form1_Load(object sender, EventArgs e)
        {
            this.reportViewer1.RefreshReport();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Multiselect = true;
            fileDialog.Title = "请选择文件";
            fileDialog.Filter = "所有excel(*.xls)|*.xls";
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    string file = fileDialog.FileName;
                    textBox1.Text = file;
                    // MessageBox.Show("已选择文件:" + file, "选择文件提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    string mystring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='";
                    mystring += textBox1.Text.ToString();
                    mystring += "';User ID=admin;Password=;Extended properties='Excel 8.0;IMEX=1;HDR=NO;'";
                    OleDbConnection cnnxls = new OleDbConnection(mystring);
                    OleDbDataAdapter myDa = new OleDbDataAdapter("Select * from [已开发票$]", cnnxls);

                    DataSet myDs = new DataSet();
                    myDa.Fill(myDs);//数据存放在myDs中了
                    DataTable dt = ChangeRows(myDs);
                    d1.DataSource = dt;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }


        private DataTable ChangeRows(DataSet ds)
        {
            DataTable dt = new DataTable();
            try
            {
                dt.Columns.Add("Dcode", typeof(string));
                dt.Columns.Add("Dtime", typeof(string));
                dt.Columns.Add("Dname", typeof(string));
                dt.Columns.Add("Dtype", typeof(string));
                dt.Columns.Add("Dnum", typeof(int));
                dt.Columns.Add("Dprice", typeof(string));
                for (int i = 1; i < ds.Tables[0].Rows.Count; i++)
                {
                    try
                    {
                        if (ds.Tables[0].Rows[i]["F2"].ToString().Trim() != "" && ds.Tables[0].Rows[i]["F3"].ToString().Trim() != "")
                        {
                            DataRow newrow = dt.NewRow();
                            newrow["Dcode"] = i.ToString();
                            newrow["Dtime"] = ToDateTimeValue(ds.Tables[0].Rows[i]["F2"].ToString());
                            newrow["Dname"] = ds.Tables[0].Rows[i]["F3"].ToString() == "" ? "" : ds.Tables[0].Rows[i]["F3"].ToString();
                            newrow["Dtype"] = ds.Tables[0].Rows[i]["F4"].ToString() == "" ? "" : ds.Tables[0].Rows[i]["F4"].ToString();
                            newrow["Dnum"] = ds.Tables[0].Rows[i]["F5"].ToString() == "" ? "" : ds.Tables[0].Rows[i]["F5"].ToString();
                            newrow["Dprice"] = ds.Tables[0].Rows[i]["F6"].ToString() == "" ? "0" : ds.Tables[0].Rows[i]["F6"].ToString();
                            dt.Rows.Add(newrow);
                        }
                    }
                    catch (Exception)
                    {
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return dt;

        }
        /// <summary>
        /// 自动打印
        /// </summary>
        /// <param name="report"></param>
        private void Export(LocalReport report)
        {
            string deviceInfo =
              "<DeviceInfo>" +
              "  <OutputFormat>EMF</OutputFormat>" +
              "  <PageWidth>8.5in</PageWidth>" +
              "  <PageHeight>11in</PageHeight>" +
              "  <MarginTop>0.15in</MarginTop>" +
              "  <MarginLeft>0.1in</MarginLeft>" +
              "  <MarginRight>0.1in</MarginRight>" +
              "  <MarginBottom>0.1in</MarginBottom>" +
              "</DeviceInfo>";
            Warning[] warnings;
            m_streams = new List<Stream>();
            try
            {
                report.Render("Image", deviceInfo, CreateStream, out warnings);//一般情况这里会出错的  使用catch得到错误原因  一般都是简单错误
            }
            catch (Exception ex)
            {
                Exception innerEx = ex.InnerException;//取内异常。因为内异常的信息才有用，才能排除问题。
                while (innerEx != null)
                {
                    //MessageBox.Show(innerEx.Message);
                    string errmessage = innerEx.Message;
                    innerEx = innerEx.InnerException;
                }
            }
            foreach (Stream stream in m_streams)
            {
                stream.Position = 0;
            }
            report.Render("Image", deviceInfo, CreateStream, out warnings);

            //foreach (Stream stream in m_streams)
            //    stream.Position = 0;
        }

        private Stream CreateStream(string name, string fileNameExtension, Encoding encoding, string mimeType, bool willSeek)
        {
            Stream stream = new FileStream(name + DateTime.Now.Millisecond + "." + fileNameExtension, FileMode.Create);
            m_streams.Add(stream);
            return stream;
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            System.Drawing.Imaging.Metafile pageImage = new System.Drawing.Imaging.Metafile(m_streams[0]);

            e.Graphics.DrawImage(pageImage, 0, 0);

            m_currentPageIndex++;
            e.HasMorePages = (m_currentPageIndex < m_streams.Count);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //生成报表
            DataTable dt1 = new DataTable();
            DataTable dt2 = new DataTable();
            DataTable dt3 = new DataTable();
            dt1.Columns.Add("Dcode", typeof(string));
            dt1.Columns.Add("Dtype", typeof(string));
            dt1.Columns.Add("Dnum", typeof(string));
            dt2.Columns.Add("Dtime", typeof(string));
            dt2.Columns.Add("Dname", typeof(string));
            dt2.Columns.Add("Dprice", typeof(string));
            dt2.Columns.Add("DSname", typeof(string));
            dt2.Columns.Add("DCname", typeof(string));
            dt2.Columns.Add("FHRTEXT", typeof(string));
            dt2.Columns.Add("SHRTEXT", typeof(string));
            dt2.Columns.Add("ISPRICE", typeof(string));
            dt3.Columns.Add("FHR", typeof(string));
            dt3.Columns.Add("CKR", typeof(string));
            dt3.Columns.Add("BName", typeof(string));
            int i = 1;
            ArrayList indexs = new ArrayList();
            float price = 0;
            foreach (DataGridViewRow row in d1.Rows)
            {
                try
                {
                    Boolean check = Convert.ToBoolean(row.Cells[0].EditedFormattedValue);
                    if (check == true)
                    {
                        if (dt2.Rows.Count == 0)
                        {
                            DataRow newrow2 = dt2.NewRow();
                            newrow2["Dtime"] = row.Cells[2].Value;
                            newrow2["Dname"] = row.Cells[3].Value;
                            newrow2["Dprice"] = row.Cells[6].Value;
                            newrow2["DSname"] = textBox2.Text.ToString();
                            newrow2["DCname"] = textBox3.Text.ToString();
                            if (radioButton1.Checked == true)
                            {
                                newrow2["FHRTEXT"] = "复核人";
                                newrow2["SHRTEXT"] = "出库人";
                                newrow2["ISPRICE"] = "0";
                            }
                            else
                            {
                                newrow2["FHRTEXT"] = "复核人";
                                newrow2["SHRTEXT"] = "入库人";
                                newrow2["ISPRICE"] = "1";
                            }
                            dt2.Rows.Add(newrow2);
                        }
                        DataRow newrow1 = dt1.NewRow();
                        newrow1["Dcode"] = i.ToString();
                        newrow1["Dtype"] = row.Cells[4].Value;
                        newrow1["Dnum"] = row.Cells[5].Value;
                        price += float.Parse(row.Cells[6].Value.ToString());
                        dt2.Rows[0]["Dprice"] = price.ToString();
                        dt1.Rows.Add(newrow1);
                        indexs.Add(row.Index);
                        i++;

                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }
            for (int n = 1; n < 5; n++)
            {
                if (dt1.Rows.Count < 5)
                {
                    DataRow newrow3 = dt1.NewRow();
                    newrow3["Dcode"] = (dt1.Rows.Count + 1).ToString();
                    newrow3["Dtype"] = "";
                    newrow3["Dnum"] = "";
                    dt1.Rows.Add(newrow3);
                }
            }
            DataRow newrow4 = dt3.NewRow();
            newrow4["FHR"] = "";
            newrow4["CKR"] = "";
            if (radioButton1.Checked == true)
            {
                newrow4["BName"] = "销售单";
            }
            else
            {
                newrow4["BName"] = "入库单";
            }
            dt3.Rows.Add(newrow4);
            int k = 0;
            for (int j = 0; j < indexs.Count; j++)
            {
                d1.Rows.RemoveAt(Convert.ToInt32(indexs[j].ToString()) - k);
                k++;
            }
            try
            {
                reportViewer1.LocalReport.ReportPath = Application.StartupPath + "\\Report1.rdlc";
                Microsoft.Reporting.WinForms.ReportDataSource rds = new Microsoft.Reporting.WinForms.ReportDataSource("PrintTab", dt1);
                Microsoft.Reporting.WinForms.ReportDataSource rds2 = new Microsoft.Reporting.WinForms.ReportDataSource("PrintTab2", dt2);
                Microsoft.Reporting.WinForms.ReportDataSource rds3 = new Microsoft.Reporting.WinForms.ReportDataSource("DataSet1", dt3);
                reportViewer1.LocalReport.DataSources.Clear();
                reportViewer1.LocalReport.DataSources.Add(rds);
                reportViewer1.LocalReport.DataSources.Add(rds2);
                reportViewer1.LocalReport.DataSources.Add(rds3);

                //reportViewer1.LocalReport.Refresh();
                reportViewer1.RefreshReport();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            //Export(reportViewer1.LocalReport);
            //m_currentPageIndex = 0;
            //printDocument1.Print();

        }
        private string ToDateTimeValue(string strNumber)
        {
            if (!string.IsNullOrEmpty(strNumber))
            {
                Decimal tempValue;
                //先检查 是不是数字;
                if (Decimal.TryParse(strNumber, out tempValue))
                {
                    //天数,取整
                    int day = Convert.ToInt32(Math.Truncate(tempValue));
                    //这里也不知道为什么. 如果是小于32,则减1,否则减2
                    //日期从1900-01-01开始累加 
                    // day = day < 32 ? day - 1 : day - 2;
                    DateTime dt = new DateTime(1900, 1, 1).AddDays(day < 32 ? (day - 1) : (day - 2));

                    //小时:减掉天数,这个数字转换小时:(* 24) 
                    Decimal hourTemp = (tempValue - day) * 24;//获取小时数
                    //取整.小时数
                    int hour = Convert.ToInt32(Math.Truncate(hourTemp));
                    //分钟:减掉小时,( * 60)
                    //这里舍入,否则取值会有1分钟误差.
                    Decimal minuteTemp = Math.Round((hourTemp - hour) * 60, 2);//获取分钟数
                    int minute = Convert.ToInt32(Math.Truncate(minuteTemp));
                    //秒:减掉分钟,( * 60)
                    //这里舍入,否则取值会有1秒误差.
                    Decimal secondTemp = Math.Round((minuteTemp - minute) * 60, 2);//获取秒数
                    int second = Convert.ToInt32(Math.Truncate(secondTemp));

                    //时间格式:00:00:00
                    string resultTimes = string.Format("{0}:{1}:{2}",
                            (hour < 10 ? ("0" + hour) : hour.ToString()),
                            (minute < 10 ? ("0" + minute) : minute.ToString()),
                            (second < 10 ? ("0" + second) : second.ToString()));

                    if (day > 0)
                        return string.Format("{0} {1}", dt.ToString("yyyy-MM-dd"), resultTimes);
                    else
                        return resultTimes;
                }
                else
                {
                    return strNumber;
                }
            }

            return string.Empty;
        }

        private void button1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
            {
                button1_Click(sender, e);
            }

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            try
            {
                // MessageBox.Show("已选择文件:" + file, "选择文件提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                string mystring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source='";
                mystring += textBox1.Text.ToString();
                mystring += "';User ID=admin;Password=;Extended properties='Excel 8.0;IMEX=1;HDR=NO;'";
                OleDbConnection cnnxls = new OleDbConnection(mystring);
                OleDbDataAdapter myDa = new OleDbDataAdapter("Select * from [已开发票$]", cnnxls);

                DataSet myDs = new DataSet();
                myDa.Fill(myDs);//数据存放在myDs中了
                DataTable dt = ChangeRows(myDs);
                d1.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            //生成报表
            DataTable dt1 = new DataTable();
            dt1.Columns.Add("Dcode", typeof(string));
            dt1.Columns.Add("Dtype", typeof(string));
            dt1.Columns.Add("Dnum", typeof(string));
            dt1.Columns.Add("Dtime", typeof(string));
            dt1.Columns.Add("Dname", typeof(string));
            dt1.Columns.Add("Dprice", typeof(string));
            dt1.Columns.Add("DSname", typeof(string));
            dt1.Columns.Add("DCname", typeof(string));
            dt1.Columns.Add("FHRTEXT", typeof(string));
            dt1.Columns.Add("SHRTEXT", typeof(string));
            dt1.Columns.Add("ISPRICE", typeof(string));
            dt1.Columns.Add("BName", typeof(string));
            ArrayList indexs = new ArrayList();
            foreach (DataGridViewRow row in d1.Rows)
            {
                try
                {
                    Boolean check = Convert.ToBoolean(row.Cells[0].EditedFormattedValue);
                    if (check == true)
                    {
                        DataRow newrow1 = dt1.NewRow();
                        newrow1["Dtime"] = row.Cells[2].Value;
                        newrow1["Dname"] = row.Cells[3].Value;
                        newrow1["Dprice"] = row.Cells[6].Value;
                        newrow1["DSname"] = textBox2.Text;
                        newrow1["DCname"] = textBox3.Text;
                        newrow1["Dcode"] = 1;
                        newrow1["Dtype"] = row.Cells[4].Value;
                        newrow1["Dnum"] = row.Cells[5].Value;
                        if (radioButton1.Checked == true)
                        {
                            newrow1["FHRTEXT"] = "复核人";
                            newrow1["SHRTEXT"] = "出库人";
                            newrow1["ISPRICE"] = "0";
                        }
                        else
                        {
                            newrow1["FHRTEXT"] = "复核人";
                            newrow1["SHRTEXT"] = "入库人";
                            newrow1["ISPRICE"] = "1";
                        }
                        if (radioButton1.Checked == true)
                        {
                            newrow1["BName"] = "销售单";
                        }
                        else
                        {
                            newrow1["BName"] = "入库单";
                        }

                        dt1.Rows.Add(newrow1);
                        indexs.Add(row.Index);
                    }
                }
                catch (Exception ex)
                {
                    continue;
                }
            }
            
            int k = 0;
            for (int j = 0; j < indexs.Count; j++)
            {
                d1.Rows.RemoveAt(Convert.ToInt32(indexs[j].ToString()) - k);
                k++;
            }
            try
            {
                reportViewer1.LocalReport.ReportPath = Application.StartupPath + "\\Report3.rdlc";
                Microsoft.Reporting.WinForms.ReportDataSource rds = new Microsoft.Reporting.WinForms.ReportDataSource("DataTable1", dt1);
                reportViewer1.LocalReport.DataSources.Clear();
                reportViewer1.LocalReport.DataSources.Add(rds);

                reportViewer1.RefreshReport();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnSelectAll_Click(object sender, EventArgs e)
        {
            //最后一行是未提交新行，不选
            for (int i = 0; i < d1.Rows.Count-1; i++)
            {
                this.d1.Rows[i].Cells[0].Value = true;
            }
        }

        private void btnNoSelectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < d1.Rows.Count; i++)
            {
                this.d1.Rows[i].Cells[0].Value = false;
            }
        }
    }
}
