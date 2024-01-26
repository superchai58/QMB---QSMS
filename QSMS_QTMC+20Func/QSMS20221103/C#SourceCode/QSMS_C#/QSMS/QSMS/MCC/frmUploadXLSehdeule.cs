using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.VisualBasic;

namespace QSMS.QSMS.MCC
{
    public partial class frmUploadXLSchedule : Form
    {

        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.MCC.UploadXLSehdeule uploadXLSehdeule = new DbLibrary.MCC.UploadXLSehdeule();
        public frmUploadXLSchedule()
        {
            InitializeComponent();
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {

            txtFilePath.Text = "";
            System.Windows.Forms.OpenFileDialog fd = new OpenFileDialog();
            fd.Filter = "Execl files (*.xlsx;*.xls)|*.xlsx;*.xls";
            string[] sheet;
            if (fd.ShowDialog() == DialogResult.OK)
            {
                txtFilePath.Text = fd.FileName;
                if (txtFilePath.Text == "") return;

                try
                {
                    sheet = pubFunction.GetExcelSheetName(txtFilePath.Text);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Read the file fail, Please make sure you close the file." + Environment.NewLine + Environment.NewLine + ex.ToString());
                    txtFilePath.Text = "";
                    return;
                }



                if (!sheet.Contains("XL_WOPlanSeq", StringComparer.OrdinalIgnoreCase))
                {
                    MessageBox.Show("The Sheet Name should be [XL_WOPlanSeq]");
                }
            }
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            DataTable dt_Excel = new DataTable();
            DataSet ds = new DataSet();

            string strFileDate;
            DateTime DateFileName;

            string strDate;
            DateTime DateExcel;

            if (txtFilePath.Text == "")
            {
                btnSelectFile.PerformClick();
                return;
            }
            strFileDate = Path.GetFileNameWithoutExtension(txtFilePath.Text);

            if (strFileDate.Split('-').Length > 1)
            {
                strFileDate = strFileDate.Split('-')[1];
                DateTime.TryParseExact(strFileDate, "yyyyMMdd", null, System.Globalization.DateTimeStyles.None, out DateFileName);

            }
            else
            {
                MessageBox.Show("File Name Error, Please select the correct file [WOPlan-YYYYMMDD]");
                return;
            }

            if (System.IO.File.Exists(txtFilePath.Text))
            {
                try
                {
                    dt_Excel = pubFunction.GetDataFromExcel(txtFilePath.Text.Trim(), "XL_WOPlanSeq");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Read the file fail, Please make sure you close the file." + Environment.NewLine + ex.ToString());
                    return;
                }

                dt_Excel.TableName = "dt_Excel";
                strDate = dt_Excel.Columns[0].ColumnName;
                DateTime.TryParse(strDate, out DateExcel);

                if (DateFileName != DateExcel)
                {
                    MessageBox.Show("XL Date is different between [File Name] and [Excel Sheet first cell]");
                    return;
                }

                dt_Excel.Columns[0].ColumnName = "Line";
                InsertDBTmp(dt_Excel, DateExcel, out ds);


            }
            else
            {
                MessageBox.Show("Please select correct file!");
                return;
            }

            if (ds.Tables[0].Rows[0]["Result"].ToString().ToUpper() != "PASS")
            {
                MessageBox.Show(ds.Tables[0].Rows[0]["Msg"].ToString(), "Upload Plan Fail");
                DataGV.DataSource = ds.Tables[1];
            }
            else
            {
                dt = uploadXLSehdeule.AssignGroupID();
                DataGV.DataSource = dt;
            }






        }

        private void InsertDBTmp(DataTable dt, DateTime date, out DataSet ds)
        {
            MemoryStream str = new MemoryStream();
            dt.WriteXml(str, true);
            str.Seek(0, SeekOrigin.Begin);
            StreamReader sr = new StreamReader(str);
            string xmlstr;
            xmlstr = sr.ReadToEnd();
            xmlstr = xmlstr.Replace("DocumentElement", "XMLData");
            xmlstr = xmlstr.Replace("dt_Excel", "XMLRows");
            xmlstr = xmlstr.ToUpper();

            string strFactory = Interaction.InputBox("Please Input Factory :F2 or F2Car", "Input Factory", "", 100, 100);
            ds = uploadXLSehdeule.UploadTemp(date.ToString("yyyyMMdd"), strFactory, xmlstr);

        }

        private void DataGV_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            if (DataGV.Columns.Contains("Flag"))
            {
                for (int i = 0; i < DataGV.Rows.Count; i++)
                {
                    if (DataGV.Rows[i].Cells["Flag"].Value.ToString() == "N") //Group is empty
                    {
                        DataGV.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;
                        DataGV.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                    }
                    if (DataGV.Rows[i].Cells["Flag"].Value.ToString() == "NEW") //New Group
                    {
                        DataGV.Rows[i].DefaultCellStyle.BackColor = Color.SkyBlue;
                    }
                    if (DataGV.Rows[i].Cells["Flag"].Value.ToString() == "Auto") //New Group
                    {
                        DataGV.Rows[i].DefaultCellStyle.BackColor = Color.GreenYellow;
                    }
                    if (DataGV.Rows[i].Cells["Flag"].Value.ToString() == "Manual") //New Group
                    {
                        DataGV.Rows[i].DefaultCellStyle.BackColor = Color.SkyBlue;
                        DataGV.Rows[i].DefaultCellStyle.ForeColor = Color.Red;
                    }
                }
            }


            if (DataGV.Columns.Contains("DBCheckMsg"))
            {
                for (int i = 0; i < DataGV.Rows.Count; i++)
                {
                    if (DataGV.Rows[i].Cells["DBCheckMsg"].Value.ToString() != "")
                    {
                        DataGV.Rows[i].DefaultCellStyle.BackColor = Color.OrangeRed;
                    }
                }
            }


            if (DataGV.Columns.Contains("WOGroup"))
            {
                for (int i = 0; i < DataGV.Rows.Count; i++)
                {
                    if (DataGV.Rows[i].Cells["WOGroup"].Value.ToString() == "")
                    {
                        DataGV.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;
                    }
                }
            }

        }

        private void DataGV_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0) return;

            string strWO, strLine, strGroup, strModel;
            if (DataGV.Columns.Contains("DBCheckMsg"))
            {
                if (DataGV.Rows[e.RowIndex].Cells["DBCheckMsg"].Value.ToString() != "")
                {
                    MessageBox.Show(DataGV.Rows[e.RowIndex].Cells["DBCheckMsg"].Value.ToString());
                    return;
                }
            }

            foreach (Control obj in gbInfo.Controls) //Clear all item
            {
                if (obj is TextBox)
                {
                    (obj as TextBox).Text = "";
                }
                else if (obj is DataGridView)
                {
                    (obj as DataGridView).DataSource = null;
                }
            }

            strWO = DataGV.Rows[e.RowIndex].Cells["WO"].Value.ToString();
            strLine = DataGV.Rows[e.RowIndex].Cells["Line"].Value.ToString();
            strGroup = DataGV.Rows[e.RowIndex].Cells["WOGroup"].Value.ToString();
            strModel = DataGV.Rows[e.RowIndex].Cells["Model"].Value.ToString();
            txtWO.Text = strWO;
            txtLine.Text = strLine;
            txtGroup.Text = strGroup;
            txtModel.Text = strModel;


        }
        private void ClearItem()
        {
            foreach (Control obj in gbInfo.Controls)
            {
                if (obj is TextBox)
                {
                    (obj as TextBox).Text = "";
                }
                else if (obj is DataGridView)
                {
                    (obj as DataGridView).DataSource = null;
                }
            }
        }
        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            DataTable dt;
            if (txtWO.Text != "")
            {
                dt = uploadXLSehdeule.GetWOAnalyze(txtWO.Text);
                gvAnalyze.DataSource = dt;
                gvAnalyze.ClearSelection();
            }
        }

        private void gvAnalyze_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            if (!gvAnalyze.Columns.Contains("Note")) return;
            for (int i = 0; i < gvAnalyze.Rows.Count; i++)
            {
                if (gvAnalyze.Rows[i].Cells["Note"].Value.ToString().IndexOf("Recommendation") != -1)
                {
                    gvAnalyze.Rows[i].DefaultCellStyle.BackColor = Color.Yellow;
                }
                else if (gvAnalyze.Rows[i].Cells["Note"].Value.ToString().IndexOf("New Group") != -1)
                {
                    gvAnalyze.Rows[i].DefaultCellStyle.BackColor = Color.Pink;
                }
            }
        }

        private void gvAnalyze_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                if (!gvAnalyze.Columns.Contains("CompareGroup")) return;

                txtNewGroup.Text = gvAnalyze.Rows[e.RowIndex].Cells["CompareGroup"].Value.ToString();
                // txtNewGroup if()
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            string msg = "";
            if (txtNewGroup.Text.Trim() != "" && txtWO.Text != "")
            {
                int iReturn;
                iReturn = uploadXLSehdeule.UpdateUpdateXL_WOPlanSeq_Tmp(txtWO.Text, txtNewGroup.Text.Trim(), ref msg);
                if (iReturn == 2)
                {
                    MessageBox.Show(msg);
                    return;
                }
                else if (iReturn == 2)
                {
                    if (MessageBox.Show(msg + Environment.NewLine + "Are you sure to continue?", "Warning", MessageBoxButtons.YesNoCancel) != DialogResult.Yes)
                    {
                        return;
                    }
                    iReturn = uploadXLSehdeule.UpdateUpdateXL_WOPlanSeq_Tmp(txtWO.Text, txtNewGroup.Text.Trim(), ref msg, "True");
                    if (iReturn == 1)
                    {
                        MessageBox.Show("Success");
                        foreach (Control obj in gbInfo.Controls) //Clear all item
                        {
                            if (obj is TextBox)
                            {
                                (obj as TextBox).Text = "";
                            }
                            else if (obj is DataGridView)
                            {
                                (obj as DataGridView).DataSource = null;
                            }
                        }
                        btnQueryTmp.PerformClick();
                    }
                    else
                    {
                        MessageBox.Show(msg);
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Success");
                    foreach (Control obj in gbInfo.Controls) //Clear all item
                    {
                        if (obj is TextBox)
                        {
                            (obj as TextBox).Text = "";
                        }
                        else if (obj is DataGridView)
                        {
                            (obj as DataGridView).DataSource = null;
                        }
                    }
                    btnQueryTmp.PerformClick();
                }
            }
        }

        private void btnGetNewGroup_Click(object sender, EventArgs e)
        {
            string tmpGroup = "";
            if (txtWO.Text == "")
            {
                return;
            }
            tmpGroup = uploadXLSehdeule.GetNewGroup(txtWO.Text);
            if (tmpGroup.Trim() == "")
            {
                MessageBox.Show("Get New Group Fail.");
            }

            txtNewGroup.Text = tmpGroup.Trim();
        }

        private void btnSaveSchedule_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            bool chkResult;
            string msg = "";
            try
            {
                chkResult = uploadXLSehdeule.CheckUpload(out msg);
                if (chkResult)
                {
                    if (msg != "")
                    {
                        if (MessageBox.Show(msg, "Warning", MessageBoxButtons.YesNoCancel) != DialogResult.Yes)
                        {
                            return;
                        }
                    }
                }
                else
                {
                    MessageBox.Show(msg, "Save Schedule Fail");
                    return;
                }
                dt = uploadXLSehdeule.SaveSchedule();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }

            DataGV.DataSource = dt;

        }

        private void frmUploadXLSchedule_Load(object sender, EventArgs e)
        {
            Type dgvType = this.DataGV.GetType();
            System.Reflection.PropertyInfo pi = dgvType.GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            pi.SetValue(this.DataGV, true, null);
        }

        private void btnReAssignGroup_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();

            try
            {
                dt = uploadXLSehdeule.ReAssignGroupID();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "Error");
            }

            DataGV.DataSource = dt;
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            string msg = "";
            bool bolResult;
            if (MessageBox.Show("Sure to delete the temp schedule?", "Warning", MessageBoxButtons.YesNo) != DialogResult.Yes)
            {
                return;
            }
            bolResult = uploadXLSehdeule.DelUploadTemp(out msg);
            if (bolResult)
            {
                MessageBox.Show("Delete OK");
            }
            else
            {
                MessageBox.Show("Delete Fail " + msg);
            }
            btnQueryTmp.PerformClick();
        }

        private void btnQueryTmp_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt = uploadXLSehdeule.GetXL_WOPlanSeq_Tmp();
            DataGV.DataSource = dt;
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            DataGV.DataSource = null;
            txtFilePath.Text = "";
            txtDate.Text = "";

            foreach (Control obj in gbInfo.Controls) //Clear all item
            {
                if (obj is TextBox)
                {
                    (obj as TextBox).Text = "";
                }
                else if (obj is DataGridView)
                {
                    (obj as DataGridView).DataSource = null;
                }
            }

        }

        private void cmsOutput_Opening(object sender, CancelEventArgs e)
        {

        }

        private void Export_Click(object sender, EventArgs e)
        {
            if (DataGV.DataSource != null)
            {
                pubFunction.CopyToExcel(DataGV, "XLWOPlan_Temp", true);
            }
        }

        private void btnOutputSchedule_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt = uploadXLSehdeule.QuerySchedule(txtDate.Text.Trim());
            if (dt.Rows.Count == 0)
            {
                MessageBox.Show("No data for the date : [" + txtDate.Text.Trim() + "]");
            }
            else
            {
                pubFunction.doExport(dt);
            }
        }

        private void frmUploadXLSchedule_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("FrmUploadXLSchedule");
        }
    }
}
