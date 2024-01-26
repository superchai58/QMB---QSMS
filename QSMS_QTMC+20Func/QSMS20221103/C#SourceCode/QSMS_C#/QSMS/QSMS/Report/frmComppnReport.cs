using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QSMS.QSMS.Report
{
    public partial class frmComppnReport : Form
    {
        DataTable dt = new DataTable();
        DataTable dt2 = new DataTable();
        DataSet ds = new DataSet();
        DbLibrary.Report.ReportProcess WorkStation = new DbLibrary.Report.ReportProcess();
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        BrLibrary.PublicFunction CopyExecl = new BrLibrary.PublicFunction();
        public frmComppnReport()
        {
            InitializeComponent();
        }

        private void frmComppnReport_Load(object sender, EventArgs e)
        {
            this.cboType.Items.Add("DifferentCompPNInfo");
        }

        private void textWO1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (this.textWO1.Text.Trim().ToString() != "" && (this.textWO1.TextLength==13 || this.textWO1.TextLength==9))
            {
                dt = WorkStation.QueryWOData(this.textWO1.Text.Trim().ToString());
                if (dt.Rows.Count != 0)
                {
                    this.textWO2.Focus();
                    return;
                }
                else
                {
                    this.textWO1.Clear();
                    this.textWO1.Focus();
                    MessageBox.Show("WO1未查询到数据！！！");
                    return;
                }
            }
        }

        private void textWO2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (this.textWO2.Text.Trim().ToString() != "" && (this.textWO2.TextLength == 13 || this.textWO2.TextLength == 9))
            {
                dt = WorkStation.QueryWOData(this.textWO2.Text.Trim().ToString());
                if (dt.Rows.Count != 0)
                {
                    this.cboType.Focus();
                    return;
                }
                else
                {
                    this.textWO2.Clear();
                    this.textWO2.Focus();
                    MessageBox.Show("WO2未查询到数据！！！");
                    return;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string WO1 = this.textWO1.Text.Trim().ToString();
            string WO2 = this.textWO2.Text.Trim().ToString();
            string Group1 = "", Group2 = "";
            if (WO1 != "" && WO2 != "" && this.cboType.Text.Trim().ToString() == "DifferentCompPNInfo")
            {
                dt = WorkStation.QueryWOData(this.textWO1.Text.Trim().ToString());
                if (dt.Rows.Count == 0)
                {
                    this.textWO1.Clear();
                    this.textWO1.Focus();
                    MessageBox.Show("WO1未查询到数据！！");
                    return;
                }
                else
                {
                    Group1 = dt.Rows[0]["Group"].ToString();
                }
                dt2 = WorkStation.QueryWOData(this.textWO2.Text.Trim().ToString());
                if (dt2.Rows.Count == 0)
                {
                    this.textWO2.Clear();
                    this.textWO2.Focus();
                    MessageBox.Show("WO2未查询到数据！！");
                    return;
                }
                else
                {
                    Group2 = dt2.Rows[0]["Group"].ToString();
                }
                if ((WO1 == WO2) || (Group1 == Group2))
                {
                    this.textWO1.Clear();
                    this.textWO2.Clear();
                    this.textWO1.Focus();
                    MessageBox.Show("它们在一个PCB组中！！！");
                    return;
                }
                else
                {
                    dt = WorkStation.GetDiffCompPNInfo(Group1, Group2);
                    if (dt.Rows.Count == 0)
                    {
                        this.textWO1.Clear();
                        this.textWO2.Clear();
                        this.textWO1.Focus();
                        MessageBox.Show("NO DATA !");
                        return;
                    }
                    else
                    {
                        CopyExecl.doExport(dt);
                        this.textWO1.Clear();
                        this.textWO2.Clear();
                        this.textWO1.Focus();
                        return;
                    }
                }
            }
            else
            {
                this.textWO1.Clear();
                this.textWO2.Clear();
                this.textWO1.Focus();
                MessageBox.Show("WO1/WO2/Type 不能为空！！！");
                return;
            }
        }

        private void frmComppnReport_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmComppnReport");
        }
    }
}
