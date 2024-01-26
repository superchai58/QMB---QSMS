using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace QSMS.QSMS.MCC
{
    public partial class frmUpLoadBom : Form
    {
        public frmUpLoadBom()
        {
            InitializeComponent();
        }
        BrLibrary.PublicFunction pubFunction = new BrLibrary.PublicFunction();
        DbLibrary.MCC.MCCProcess MCC = new DbLibrary.MCC.MCCProcess();
        private void frmUpLoadBom_FormClosed(object sender, FormClosedEventArgs e)
        {
            pubFunction.RemoveForm("frmUpLoadBom");
        }

        private void frmUpLoadBom_Load(object sender, EventArgs e)
        {
            DataTable dt = MCC.CheckUploadReplacePNRight();
            if (dt.Rows.Count > 0)
            {
                CboFuncType.Items.Add("REPLACEPN");
            }
            CboFuncType.Items.Add("QSMS_MEBom");
            CboFuncType.Items.Add("CastRate");
            CboFuncType.Items.Add("ComppnInSpectRule");
            CboFuncType.Items.Add("Component_Data");
            CboFuncType.Items.Add("FujiBrdSeqMapping");
            CboFuncType.Items.Add("MaterialToWHID");
            CboFuncType.Items.Add("NoMachineDropCompPN");
            CboFuncType.Items.Add("OneByOne");
            CboFuncType.Items.Add("SingleSideBrd");
            CboFuncType.Items.Add("TraySlot");
            CboFuncType.Items.Add("UNCHKCOMP");
            CboFuncType.Items.Add("XL_WOPlanSeq");
            CboFuncType.Items.Add("XL_WOPlanLine");
            CboFuncType.Items.Add("XL_ImplementPN");
            CboFuncType.Items.Add("XL_PNOneByOne");
            CboFuncType.Items.Add("XL_DoubleTables");
            CboFuncType.Items.Add("XL_MaxDIDMaintainQty");
            CboFuncType.Items.Add("NOCheckReplacePNSplicing");
            CboFuncType.Items.Add("Machine_Data");
            btnExcel.Enabled = false;
        }

        private void CboFuncType_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(CboFuncType.Text.Trim() != "")
            {
                btnExcel.Enabled = true;
            }
            switch(CboFuncType.Text.ToUpper())
            {               
                case "QSMS_MEBOM":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\QSMS_MEBOM.xls";
                    break;
                case "SINGLESIDEBRD":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\SINGLESIDEBRD.xls";
                    break;
                case "REPLACEPN":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\REPLACEPN.xls";
                    break;
                case "UNCHKCOMP":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\UNCHKCOMP.xls";
                    break;
                case "FUJIBRDSEQMAPPING":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\FujiBrdSeqMapping.xls";
                    break;
                case "TRAYSLOT":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\TraySlot.xls";
                    break;
                case "CASTRATE":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\CASTRATE.xls";
                    break;
                case "ONEBYONE":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\ONEBYONE.xls";
                    break;
                case "NOMACHINEDROPCOMPPN":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\NOMACHINEDROPCOMPPN.xls";
                    break;
                case "COMPPNINSPECTRULE":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\COMPPNINSPECTRULE.xls";
                    break;
                case "XL_WOPLANSEQ":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\XL_WOPLANSEQ.xls";
                    break;
                case "XL_WOPLANLINE":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\XL_WOPlanLine.xls";
                    break;
                case "XL_IMPLEMENTPN":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\XL_IMPLEMENTPN.xls";
                    break;
                case "MATERIALTOWHID":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\MATERIALTOWHID.xls";
                    break;
                case "XL_PNONEBYONE":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\XL_PNONEBYONE.xls";
                    break;
                case "XL_DOUBLETABLES":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\XL_DOUBLETABLES.xls";
                    break;
                case "XL_MAXDIDMAINTAINQTY":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\XL_MAXDIDMAINTAINQTY.xls";
                    break;
                case "NOCHECKREPLACEPNSPLICING":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\NOCheckReplacePNSplicing.xls";
                    break;              
                case "COMPONENT_DATA":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\Component_data.xls";
                    break;
                case "MACHINE_DATA":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\Machine_DATA.xls";
                    break;
                case "":
                    btnExcel.Enabled = false;
                    lblFileFormat.Text ="";
                    break;
                default:
                    btnExcel.Enabled = false;
                    MessageBox.Show("Please check the right sheet name.");
                    break;
            }
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            try
            {
                DataTable dt;
                if (CboFuncType.Text.Trim() == "")
                {
                    MessageBox.Show("Please check the right sheet name.");
                    return;
                }
                else
                {
                    dt = MCC.GetDataByFuncType(CboFuncType.Text.Trim());
                    if (dt.Rows.Count > 0)
                    {
                        if (Parameter.chkDomain == "N")
                        {

                        }
                        else
                        {
                            pubFunction.doExport(dt);
                            MessageBox.Show("Data Export Complete!");
                        }
                    }
                    else
                    {
                        MessageBox.Show("No Data!");
                        return;
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void btnSelectFile_Click(object sender, EventArgs e)
        {
            txtFilePath.Text = "";
            System.Windows.Forms.OpenFileDialog fd = new OpenFileDialog();
            fd.Filter = "Execl files (*.xlsx;*.xls)|*.xlsx;*.xls";
            if (fd.ShowDialog() == DialogResult.OK)
            {
                txtFilePath.Text = fd.FileName;
                if (txtFilePath.Text == "") return;


                string[] sheet = pubFunction.GetExcelSheetName(txtFilePath.Text);
                cboSheetName.Text = "";
                cboSheetName.Items.Clear();
                pubFunction.BindComboBox(cboSheetName, sheet);
            }
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            if(cboSheetName.Text.Trim() == "" || txtFilePath.Text.Trim() == "" || CboFuncType.Text.Trim() == "")
            {
                MessageBox.Show("SheetName and FilePath and FuncType can not be NULL !", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if ((CboFuncType.Text.Trim() != "FujiBrdSeqMapping" && CboFuncType.Text.Trim() != "PhilipsBrdSeqMapping") && cboSheetName.Text.Trim().ToUpper() != CboFuncType.Text.Trim().ToUpper())
            {
                MessageBox.Show("Function type does not match the sheet name ,please check! And sheet name should the same as function type", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            switch(CboFuncType.Text.Trim().ToUpper())
            {
                case "QSMS_MEBOM":
                    Load_QSMS_BOM();
                    break;
                case "SINGLESIDEBRD":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\SINGLESIDEBRD.xls";
                    break;
                case "REPLACEPN":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\REPLACEPN.xls";
                    break;
                case "UNCHKCOMP":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\UNCHKCOMP.xls";
                    break;
                case "FUJIBRDSEQMAPPING":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\FujiBrdSeqMapping.xls";
                    break;
                case "TRAYSLOT":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\TraySlot.xls";
                    break;
                case "CASTRATE":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\CASTRATE.xls";
                    break;
                case "ONEBYONE":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\ONEBYONE.xls";
                    break;
                case "NOMACHINEDROPCOMPPN":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\NOMACHINEDROPCOMPPN.xls";
                    break;
                case "COMPPNINSPECTRULE":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\COMPPNINSPECTRULE.xls";
                    break;
                case "XL_WOPLANSEQ":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\XL_WOPLANSEQ.xls";
                    break;
                case "XL_WOPLANLINE":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\XL_WOPlanLine.xls";
                    break;
                case "XL_IMPLEMENTPN":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\XL_IMPLEMENTPN.xls";
                    break;
                case "MATERIALTOWHID":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\MATERIALTOWHID.xls";
                    break;
                case "XL_PNONEBYONE":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\XL_PNONEBYONE.xls";
                    break;
                case "XL_DOUBLETABLES":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\XL_DOUBLETABLES.xls";
                    break;
                case "XL_MAXDIDMAINTAINQTY":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\XL_MAXDIDMAINTAINQTY.xls";
                    break;
                case "NOCHECKREPLACEPNSPLICING":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\NOCheckReplacePNSplicing.xls";
                    break;
                case "COMPONENT_DATA":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\Component_data.xls";
                    break;
                case "MACHINE_DATA":
                    lblFileFormat.Text = Application.StartupPath.Trim() + "\\Template\\Machine_DATA.xls";
                    break;
                case "":
                    btnExcel.Enabled = false;
                    lblFileFormat.Text = "";
                    break;
                default:
                    btnExcel.Enabled = false;
                    MessageBox.Show("Please check the right sheet name.");
                    break;
            }
        }
        private void Load_QSMS_BOM()
        {
            string Factory, Line, Machine, Jobpn, Version, Slot, compPN, Qty, LR, jobgroup, BuildType, Side, location;

            if (txtFilePath.Text.Trim() == "" || cboSheetName.Text.Trim() == "")
            {
                MessageBox.Show("Please Choose the File & Sheet!");
                return;
            }

            DataTable ds = pubFunction.GetDataFromExcel(txtFilePath.Text.Trim(),cboSheetName.Text.Trim());
            if(ds.Rows.Count <=0)
            {
                MessageBox.Show("No Data Can Be Used Upload!");
                return;
            }
            for (int i = 0; i < ds.Rows.Count; i++)
            {
                Machine = ds.Rows[i][0].ToString().Trim();
                Jobpn = ds.Rows[i][1].ToString().Trim();
                Version = ds.Rows[i][2].ToString().Trim();
                Slot = ds.Rows[i][3].ToString().Trim();
                compPN = ds.Rows[i][4].ToString().Trim();
                Qty = ds.Rows[i][5].ToString().Trim();
                LR = ds.Rows[i][6].ToString().Trim();
                jobgroup = ds.Rows[i][7].ToString().Trim();
                BuildType = ds.Rows[i][8].ToString().Trim();
                Side = ds.Rows[i][9].ToString().Trim();
                Factory = ds.Rows[i][10].ToString().Trim();
            }
        }
    }
}
