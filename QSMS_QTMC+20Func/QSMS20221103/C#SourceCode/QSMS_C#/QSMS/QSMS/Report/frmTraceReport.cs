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
    public partial class frmTraceReport : Form
    {
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DataTable dt = new DataTable();
        DbLibrary.Report.ReportProcess ReportProcess = new DbLibrary.Report.ReportProcess();
        public frmTraceReport()
        {
            InitializeComponent();
        }

        private void frmTraceReport_Load(object sender, EventArgs e)
        {
            CbbDataType.Items.Clear();
            CbbDataType.Items.Add("Trace By SN");
            CbbDataType.Items.Add("Trace By WO");
            CbbDataType.Items.Add("Trace By DID");
            CbbDataType.Items.Add("Trace By CompPN");
            DTPBeginTime.Text = "080000";
            DTPEndTime.Text = "080000";
        }

        private void cmdGetData_Click(object sender, EventArgs e)
        {
            try
            {
                string resultType, strDestPath;
                labelInfor.Text = "";
                lblInfor.Text = "";
                string BeginDate = Convert.ToDateTime(DTPBeginDate.Text).ToString("yyyyMMdd");
                string EndDate = Convert.ToDateTime(DTPEndDate.Text).ToString("yyyyMMdd");
                string BeginDateTime = BeginDate + DTPBeginTime.Text.ToString().Trim();
                string EndDateTime = EndDate + DTPEndTime.Text.ToString().Trim();
                if (CbbDataType.Text == null)
                {
                    MessageBox.Show("Please select the Data Type!");
                    return;
                }
                if (optExcel.Checked == true)
                {
                    resultType = "EXCEL";
                }
                else
                {
                    resultType = "TEXT";
                }
                strDestPath = pubFunction.ConfigListGetValue("TraceReportPath");
                if (!System.IO.Directory.Exists(strDestPath))
                {
                    System.IO.Directory.CreateDirectory(strDestPath);
                }
                if (CbbDataType.Text.ToString().Trim().ToUpper() == "TRACE BY COMPPN")
                {
                    if (Convert.ToInt64(BeginDateTime) > Convert.ToInt64(EndDateTime))
                    {
                        MessageBox.Show("Please input right start date time and end date time!");
                        return;
                    }
                    if (TxtCompPN.Text.ToString().Trim() == null)
                    {
                        MessageBox.Show("Please input the CompPN!");
                        TxtCompPN.Focus();
                        return;
                    }
                    if (Convert.ToInt64(BeginDate) - Convert.ToInt64(EndDate) > 31)
                    {
                        MessageBox.Show("系统一次查询的时间跨度不能超过31天!请分多次查询!");
                        return;
                    }
                    dt = ReportProcess.GetSNByComp(TxtCompPN.Text.ToString().Trim(), txtVendorCode.Text.ToString().Trim(), txtDateCode.Text.ToString().Trim(), txtLotCode.Text.ToString().Trim(), BeginDateTime, EndDateTime, txtModel.Text.ToString().Trim(), resultType);
                }
                else
                {
                    if (OptSN.Checked == true)
                    {
                        if (TxtSN.Text.ToString().Trim() == null)
                        {
                            MessageBox.Show("Please input the SN or DID or WO!");
                            return;
                        }
                        if (CbbDataType.Text.ToString().Trim().ToUpper() == "TRACE BY SN")
                        {
                            dt = ReportProcess.TraceReport_GetCompBySN("one", TxtSN.Text.ToString().Trim(), TxtCompPN.ToString().Trim(), resultType);
                        }
                        if (CbbDataType.Text.ToString().Trim().ToUpper() == "TRACE BY DID")
                        {
                            dt = ReportProcess.TraceReport_GetSNByDID("one", TxtSN.Text.ToString().Trim(), resultType);
                        }
                        if (CbbDataType.Text.ToString().Trim().ToUpper() == "TRACE BY WO")
                        {
                            dt = ReportProcess.TraceReport_GetCompByWO(TxtSN.Text.ToString().Trim(), TxtCompPN.Text.ToString().Trim(), resultType);
                        }
                    }
                    else
                    {
                        if (CbbDataType.Text.ToString().Trim().ToUpper() == "TRACE BY SN")
                        {
                            dt = ReportProcess.TraceReport_GetCompBySN("Batch", "", "", resultType);
                        }
                        if (CbbDataType.Text.ToString().Trim().ToUpper() == "TRACE BY DID")
                        {
                            dt = ReportProcess.TraceReport_GetSNByDID("Batch", "", resultType);
                        }
                        if (CbbDataType.Text.ToString().Trim().ToUpper() == "TRACE BY WO")
                        {
                            dt = ReportProcess.TraceReport_GetCompByWO("", "", resultType);
                        }
                    }
                }
                if (resultType == "EXCEL")
                {
                    pubFunction.doExportSave(dt);
                }
                else
                {
                    strDestPath = strDestPath + "\\" + dt.Rows[0]["FileName"].ToString().Trim();
                    lblInfor.Text = "从服务器:" + dt.Rows[0]["FileName"].ToString().Trim() + " 复制查询的 TraceReport 数据到本地目录:" + strDestPath + ",请核对!";
                    System.IO.File.Copy(dt.Rows[0]["FilePath"].ToString().Trim(), strDestPath, true);
                    MessageBox.Show("从服务器:" + dt.Rows[0]["FileName"].ToString().Trim() + " 复制查询的 TraceReport 数据到本地目录:" + strDestPath + "完成,请核对!");
                    string strDeleteFile = "D:\\TraceReportData\\" + dt.Rows[0]["FileName"].ToString().Trim();
                    dt = ReportProcess.TraceReport_DeleteFile(strDeleteFile);
                }
                MessageBox.Show("Data create OK!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        //private void CbbDataType_Click(object sender, EventArgs e)
        //{
           
        //}

        private void CMDChosefile_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "请选择要上传的文件";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Txtpath.Text = dialog.FileName;
            }
        }

        private void inputSN_Click(object sender, EventArgs e)
        {
            try
            {
                string SN = "", CompPN = "";
                if (Txtpath.Text.Trim() == "")
                {
                    MessageBox.Show("上传路劲不能为空！");
                    return;
                }

                ReportProcess.DeleteTempSN();

                dt = ImportExcelToDataTable2(Txtpath.Text.Trim());
                this.dtData.DataSource = dt;
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    SN = dt.Rows[i][0].ToString().Trim();
                    CompPN = dt.Rows[i][1].ToString().Trim();
                    ReportProcess.TraceReport_TempSN(SN, CompPN);
                }
                dt = ReportProcess.QueryTempSN();
                DataGridSN.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }
        private DataTable ImportExcelToDataTable2(string path)
        {
            string conStr = string.Format("Provider=Microsoft.ACE.OLEDB.12.0; Data source={0}; Extended Properties=Excel 12.0;", path);
            using (System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection(conStr))
            {
                conn.Open();
                //获取所有Sheet的相关信息
                DataTable dtSheet = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, null);
                //获取第一个 Sheet的名称
                string sheetName = dtSheet.Rows[0]["Table_Name"].ToString();
                string sql = string.Format("select * from [{0}]", sheetName);
                using (System.Data.OleDb.OleDbDataAdapter oda = new System.Data.OleDb.OleDbDataAdapter(sql, conn))
                {
                    DataTable dt = new DataTable();
                    oda.Fill(dt);
                    return dt;
                }
            }
        }

        private void CbbDataType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CbbDataType.Text.ToString().Trim() == "Trace By SN")
            {
                lblSNWO.Text = "SN";
            }
            else if (CbbDataType.Text.ToString().Trim() == "Trace By WO")
            {
                lblSNWO.Text = "WO";
            }
            else if (CbbDataType.Text.ToString().Trim() == "Trace By DID")
            {
                lblSNWO.Text = "DID";
            }
        }
        private void DataGridSN_MouseClick(object sender, MouseEventArgs e)
        {
            for (int i = 0; i < DataGridSN.Rows.Count; i++)
            {
                for (int y = 0; y < DataGridSN.ColumnCount; y++)
                {
                    if (DataGridSN.Rows[i].Cells[y].Selected == true)
                    {
                        TxtSN.Text = DataGridSN.Rows[i].Cells[0].Value.ToString();
                        TxtCompPN.Text = DataGridSN.Rows[i].Cells[1].Value.ToString();
                        return;
                    }
                }
            }
        }

        private void frmTraceReport_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmTraceReport");
        }
    }
}
