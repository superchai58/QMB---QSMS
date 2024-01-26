using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QSMS.QSMS.QMS
{
    public partial class frmSetInterDIO : Form
    {
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.QMS.SetInterDIOProcess SetInterDIO = new DbLibrary.QMS.SetInterDIOProcess();
        CheckBox[] chkMachine = new CheckBox[48];
        int cQty = 0, ClickFlg=0;
        public frmSetInterDIO()
        {
            InitializeComponent();
        }

        private void frmSetInterDIO_Load(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            DataTable dt1 = new DataTable();
            dt1 = SetInterDIO.GetMachine1();
            dt = SetInterDIO.GetMachine();
            if (dt.Rows.Count == dt.Rows.Count)
            {
                chkAllLine.Checked = true;
            }
            GetLine();
            for (int i = 0; i < 48; i++)
            {

                chkMachine[i] = new CheckBox();
                chkMachine[i].Name = "chkMachine" + Convert.ToString(i);
                chkMachine[i].Text = "Check2";
                chkMachine[i].ForeColor = Color.Black;
                chkMachine[i].Location = new Point(3 + (i % 8) * 140, 40 + (i / 8) * 30);
                //chkMachine[i].Size = new Size(90, 25);
                chkMachine[i].Size = new Size(120, 25);
                chkMachine[i].BackColor = Color.Gainsboro;
                chkMachine[i].Font = new System.Drawing.Font("微软雅黑", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
                //lblSlot[(j - 1) * MaxSlot + k + 1].BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
                chkMachine[i].TextAlign = ContentAlignment.MiddleCenter;
                //chkMachine[(j - 1) * MaxSlot + k + 1].DoubleClick += new EventHandler(chkMachine_DoubleClick);
                fraMachine.Controls.Add(chkMachine[i]);
            }
            fraMachine.Visible = false;

        }
        public void GetLine()
        {
            DataTable dt = new DataTable();
            dt = SetInterDIO.GetLine();
            if (dt.Rows.Count > 0)
            {
                CboLine.Items.Clear();
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    CboLine.Items.Add(dt.Rows[i]["Line"].ToString());
                }

            }
        }
        public void GetSet()
        {
            DataTable dt = new DataTable();
            dt = SetInterDIO.GetSet(CboLine.Text);
            if (dt.Rows.Count > 0)
            {
                fraMachine.Visible = true;
                if (dt.Rows.Count < 48)
                {
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {
                        chkMachine[i].Visible = true;
                        chkMachine[i].Text = dt.Rows[i]["Machine"].ToString();
                        if (dt.Rows[i]["DisableInterlock"].ToString() == "1")
                        {
                            chkMachine[i].Checked = true;
                        }
                        else
                        {
                            chkMachine[i].Checked = false;
                        }
                    }
                }
                else if (dt.Rows.Count > 48)
                {
                    for (int j = 0; j < 48; j++)
                    {
                        chkMachine[j].Visible = true;
                        chkMachine[j].Text = dt.Rows[j]["Machine"].ToString();
                        if (dt.Rows[j]["DisableInterlock"].ToString() == "1")
                        {
                            chkMachine[j].Checked = true;
                        }
                        else
                        {
                            chkMachine[j].Checked = false;
                        }
                    }
                }
                else
                {
                    return;
                }
                cQty = dt.Rows.Count;
            }

            //    for (int i = 0; i < 48; i++)
            //    {
            //        if (dt.Rows.Count < 48)
            //        {
            //            chkMachine[i].Visible = true;
            //            chkMachine[i].Text = dt.Rows[i]["Machine"].ToString();
            //            if (dt.Rows[i]["DisableInterlock"].ToString() == "1")
            //            {
            //                chkMachine[i].Checked = true;
            //            }
            //            else
            //            {
            //                chkMachine[i].Checked = false;
            //            }
            //        }
            //        else
            //        {
            //            chkMachine[i].Visible = false;
            //        }
            //    }
            //    cQty = dt.Rows.Count;
            //}
        }

        private void cmdExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void chkAllLine_Click(object sender, EventArgs e)
        {
            ClickFlg = 1;
        }

        private void cmdOK_Click(object sender, EventArgs e)
        {
            string chkAllLineData = "", chkMachineChecked="";
            if (ClickFlg == 1)
            {
                if (chkAllLine.Checked == true)
                {
                    chkAllLineData = "1";
                }
                else
                {
                    chkAllLineData = "0";
                }
                SetInterDIO.SetInterDIO(ClickFlg, chkAllLineData, Parameter.UID, "", "");
            }
            else if (ClickFlg == 2)
            {
                for(int i=0;i<cQty;i++)
                {
                    if (chkMachine[i].Checked == true)
                    {
                        chkMachineChecked = "1";
                    }
                    else
                    {
                        chkMachineChecked = "0";
                    }
                    SetInterDIO.SetInterDIO(ClickFlg, chkAllLineData, Parameter.UID, chkMachineChecked, chkMachine[i].Text);
                }
            }
        }

        private void CboLine_Click(object sender, EventArgs e)
        {
            //GetSet();
            //fraMachine.Visible = true;
            //ClickFlg = 2;
        }

        private void frmSetInterDIO_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmSetInterDIO");
        }

        private void CboLine_SelectedValueChanged(object sender, EventArgs e)
        {
            GetSet();
            fraMachine.Visible = true;
            ClickFlg = 2;
        }
    }
}
