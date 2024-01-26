namespace QSMS.QSMS.MCC
{
    partial class frmDIDInteGration
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lblFactory = new System.Windows.Forms.Label();
            this.cboFactory = new System.Windows.Forms.ComboBox();
            this.txtQty = new System.Windows.Forms.TextBox();
            this.lblQty = new System.Windows.Forms.Label();
            this.lblSide = new System.Windows.Forms.Label();
            this.cboSide = new System.Windows.Forms.ComboBox();
            this.lblLine = new System.Windows.Forms.Label();
            this.cboLine = new System.Windows.Forms.ComboBox();
            this.txtEndT = new System.Windows.Forms.TextBox();
            this.txtBeginT = new System.Windows.Forms.TextBox();
            this.DTPEndDate = new System.Windows.Forms.DateTimePicker();
            this.lblEndT = new System.Windows.Forms.Label();
            this.DTPBeginDate = new System.Windows.Forms.DateTimePicker();
            this.lblBeginT = new System.Windows.Forms.Label();
            this.txtVendorCode = new System.Windows.Forms.TextBox();
            this.lblVendorCode = new System.Windows.Forms.Label();
            this.txtLotCode = new System.Windows.Forms.TextBox();
            this.lblLotCode = new System.Windows.Forms.Label();
            this.txtCompPN = new System.Windows.Forms.TextBox();
            this.lblCompPN = new System.Windows.Forms.Label();
            this.txtDateCode = new System.Windows.Forms.TextBox();
            this.lblDateCode = new System.Windows.Forms.Label();
            this.txtDID = new System.Windows.Forms.TextBox();
            this.lblDID = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupboxLabel = new System.Windows.Forms.GroupBox();
            this.OptOld = new System.Windows.Forms.RadioButton();
            this.OptNew = new System.Windows.Forms.RadioButton();
            this.btnExit = new System.Windows.Forms.Button();
            this.btnQuery = new System.Windows.Forms.Button();
            this.btnReprint = new System.Windows.Forms.Button();
            this.btnExcel = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnRefresh = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.DG_Result = new System.Windows.Forms.DataGridView();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.lblMsg = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupboxLabel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DG_Result)).BeginInit();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.lblFactory);
            this.groupBox1.Controls.Add(this.cboFactory);
            this.groupBox1.Controls.Add(this.txtQty);
            this.groupBox1.Controls.Add(this.lblQty);
            this.groupBox1.Controls.Add(this.lblSide);
            this.groupBox1.Controls.Add(this.cboSide);
            this.groupBox1.Controls.Add(this.lblLine);
            this.groupBox1.Controls.Add(this.cboLine);
            this.groupBox1.Controls.Add(this.txtEndT);
            this.groupBox1.Controls.Add(this.txtBeginT);
            this.groupBox1.Controls.Add(this.DTPEndDate);
            this.groupBox1.Controls.Add(this.lblEndT);
            this.groupBox1.Controls.Add(this.DTPBeginDate);
            this.groupBox1.Controls.Add(this.lblBeginT);
            this.groupBox1.Controls.Add(this.txtVendorCode);
            this.groupBox1.Controls.Add(this.lblVendorCode);
            this.groupBox1.Controls.Add(this.txtLotCode);
            this.groupBox1.Controls.Add(this.lblLotCode);
            this.groupBox1.Controls.Add(this.txtCompPN);
            this.groupBox1.Controls.Add(this.lblCompPN);
            this.groupBox1.Controls.Add(this.txtDateCode);
            this.groupBox1.Controls.Add(this.lblDateCode);
            this.groupBox1.Controls.Add(this.txtDID);
            this.groupBox1.Controls.Add(this.lblDID);
            this.groupBox1.Location = new System.Drawing.Point(9, 5);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Size = new System.Drawing.Size(724, 193);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Info";
            // 
            // lblFactory
            // 
            this.lblFactory.BackColor = System.Drawing.SystemColors.Info;
            this.lblFactory.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblFactory.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblFactory.Location = new System.Drawing.Point(545, 18);
            this.lblFactory.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblFactory.Name = "lblFactory";
            this.lblFactory.Size = new System.Drawing.Size(77, 24);
            this.lblFactory.TabIndex = 32;
            this.lblFactory.Text = "厂区:";
            this.lblFactory.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // cboFactory
            // 
            this.cboFactory.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cboFactory.FormattingEnabled = true;
            this.cboFactory.Location = new System.Drawing.Point(627, 17);
            this.cboFactory.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.cboFactory.Name = "cboFactory";
            this.cboFactory.Size = new System.Drawing.Size(84, 24);
            this.cboFactory.TabIndex = 33;
            // 
            // txtQty
            // 
            this.txtQty.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.txtQty.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtQty.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtQty.Location = new System.Drawing.Point(627, 54);
            this.txtQty.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtQty.Name = "txtQty";
            this.txtQty.Size = new System.Drawing.Size(83, 26);
            this.txtQty.TabIndex = 31;
            // 
            // lblQty
            // 
            this.lblQty.BackColor = System.Drawing.SystemColors.Info;
            this.lblQty.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblQty.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblQty.Location = new System.Drawing.Point(545, 55);
            this.lblQty.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblQty.Name = "lblQty";
            this.lblQty.Size = new System.Drawing.Size(77, 24);
            this.lblQty.TabIndex = 30;
            this.lblQty.Text = "数量:";
            this.lblQty.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblSide
            // 
            this.lblSide.BackColor = System.Drawing.SystemColors.Info;
            this.lblSide.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblSide.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblSide.Location = new System.Drawing.Point(386, 54);
            this.lblSide.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblSide.Name = "lblSide";
            this.lblSide.Size = new System.Drawing.Size(57, 24);
            this.lblSide.TabIndex = 26;
            this.lblSide.Text = "面别:";
            this.lblSide.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // cboSide
            // 
            this.cboSide.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cboSide.FormattingEnabled = true;
            this.cboSide.Location = new System.Drawing.Point(448, 54);
            this.cboSide.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.cboSide.Name = "cboSide";
            this.cboSide.Size = new System.Drawing.Size(84, 24);
            this.cboSide.TabIndex = 27;
            // 
            // lblLine
            // 
            this.lblLine.BackColor = System.Drawing.SystemColors.Info;
            this.lblLine.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblLine.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblLine.Location = new System.Drawing.Point(386, 17);
            this.lblLine.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblLine.Name = "lblLine";
            this.lblLine.Size = new System.Drawing.Size(57, 24);
            this.lblLine.TabIndex = 24;
            this.lblLine.Text = "线别:";
            this.lblLine.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // cboLine
            // 
            this.cboLine.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cboLine.FormattingEnabled = true;
            this.cboLine.Location = new System.Drawing.Point(448, 18);
            this.cboLine.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.cboLine.Name = "cboLine";
            this.cboLine.Size = new System.Drawing.Size(84, 24);
            this.cboLine.TabIndex = 25;
            // 
            // txtEndT
            // 
            this.txtEndT.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtEndT.Location = new System.Drawing.Point(223, 53);
            this.txtEndT.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtEndT.Name = "txtEndT";
            this.txtEndT.Size = new System.Drawing.Size(98, 26);
            this.txtEndT.TabIndex = 23;
            // 
            // txtBeginT
            // 
            this.txtBeginT.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtBeginT.Location = new System.Drawing.Point(223, 16);
            this.txtBeginT.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtBeginT.Name = "txtBeginT";
            this.txtBeginT.Size = new System.Drawing.Size(98, 26);
            this.txtBeginT.TabIndex = 20;
            // 
            // DTPEndDate
            // 
            this.DTPEndDate.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.DTPEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.DTPEndDate.Location = new System.Drawing.Point(94, 54);
            this.DTPEndDate.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.DTPEndDate.Name = "DTPEndDate";
            this.DTPEndDate.Size = new System.Drawing.Size(125, 26);
            this.DTPEndDate.TabIndex = 22;
            this.DTPEndDate.Value = new System.DateTime(2019, 11, 22, 12, 41, 44, 0);
            // 
            // lblEndT
            // 
            this.lblEndT.BackColor = System.Drawing.SystemColors.Info;
            this.lblEndT.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblEndT.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblEndT.Location = new System.Drawing.Point(4, 54);
            this.lblEndT.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblEndT.Name = "lblEndT";
            this.lblEndT.Size = new System.Drawing.Size(86, 24);
            this.lblEndT.TabIndex = 21;
            this.lblEndT.Text = "结束时间:";
            this.lblEndT.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // DTPBeginDate
            // 
            this.DTPBeginDate.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.DTPBeginDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.DTPBeginDate.Location = new System.Drawing.Point(94, 16);
            this.DTPBeginDate.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.DTPBeginDate.Name = "DTPBeginDate";
            this.DTPBeginDate.Size = new System.Drawing.Size(125, 26);
            this.DTPBeginDate.TabIndex = 19;
            this.DTPBeginDate.Value = new System.DateTime(2019, 11, 22, 12, 41, 44, 0);
            // 
            // lblBeginT
            // 
            this.lblBeginT.BackColor = System.Drawing.SystemColors.Info;
            this.lblBeginT.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblBeginT.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblBeginT.Location = new System.Drawing.Point(4, 17);
            this.lblBeginT.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblBeginT.Name = "lblBeginT";
            this.lblBeginT.Size = new System.Drawing.Size(86, 24);
            this.lblBeginT.TabIndex = 18;
            this.lblBeginT.Text = "开始时间:";
            this.lblBeginT.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtVendorCode
            // 
            this.txtVendorCode.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.txtVendorCode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtVendorCode.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtVendorCode.Location = new System.Drawing.Point(489, 89);
            this.txtVendorCode.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtVendorCode.Name = "txtVendorCode";
            this.txtVendorCode.Size = new System.Drawing.Size(221, 26);
            this.txtVendorCode.TabIndex = 5;
            // 
            // lblVendorCode
            // 
            this.lblVendorCode.BackColor = System.Drawing.SystemColors.Info;
            this.lblVendorCode.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblVendorCode.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblVendorCode.Location = new System.Drawing.Point(386, 90);
            this.lblVendorCode.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblVendorCode.Name = "lblVendorCode";
            this.lblVendorCode.Size = new System.Drawing.Size(98, 24);
            this.lblVendorCode.TabIndex = 4;
            this.lblVendorCode.Text = "厂商代码:";
            this.lblVendorCode.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtLotCode
            // 
            this.txtLotCode.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.txtLotCode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtLotCode.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtLotCode.Location = new System.Drawing.Point(489, 124);
            this.txtLotCode.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtLotCode.Name = "txtLotCode";
            this.txtLotCode.Size = new System.Drawing.Size(221, 26);
            this.txtLotCode.TabIndex = 9;
            // 
            // lblLotCode
            // 
            this.lblLotCode.BackColor = System.Drawing.SystemColors.Info;
            this.lblLotCode.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblLotCode.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblLotCode.Location = new System.Drawing.Point(386, 125);
            this.lblLotCode.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblLotCode.Name = "lblLotCode";
            this.lblLotCode.Size = new System.Drawing.Size(98, 24);
            this.lblLotCode.TabIndex = 8;
            this.lblLotCode.Text = "生产批号:";
            this.lblLotCode.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtCompPN
            // 
            this.txtCompPN.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.txtCompPN.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtCompPN.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtCompPN.Location = new System.Drawing.Point(94, 89);
            this.txtCompPN.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtCompPN.Name = "txtCompPN";
            this.txtCompPN.Size = new System.Drawing.Size(226, 26);
            this.txtCompPN.TabIndex = 3;
            // 
            // lblCompPN
            // 
            this.lblCompPN.BackColor = System.Drawing.SystemColors.Info;
            this.lblCompPN.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblCompPN.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblCompPN.Location = new System.Drawing.Point(4, 89);
            this.lblCompPN.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblCompPN.Name = "lblCompPN";
            this.lblCompPN.Size = new System.Drawing.Size(86, 24);
            this.lblCompPN.TabIndex = 2;
            this.lblCompPN.Text = "料号:";
            this.lblCompPN.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtDateCode
            // 
            this.txtDateCode.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
            this.txtDateCode.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtDateCode.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtDateCode.Location = new System.Drawing.Point(94, 124);
            this.txtDateCode.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtDateCode.Name = "txtDateCode";
            this.txtDateCode.Size = new System.Drawing.Size(226, 26);
            this.txtDateCode.TabIndex = 7;
            // 
            // lblDateCode
            // 
            this.lblDateCode.BackColor = System.Drawing.SystemColors.Info;
            this.lblDateCode.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblDateCode.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblDateCode.Location = new System.Drawing.Point(4, 125);
            this.lblDateCode.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblDateCode.Name = "lblDateCode";
            this.lblDateCode.Size = new System.Drawing.Size(86, 24);
            this.lblDateCode.TabIndex = 6;
            this.lblDateCode.Text = "生产日期:";
            this.lblDateCode.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtDID
            // 
            this.txtDID.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtDID.Location = new System.Drawing.Point(94, 160);
            this.txtDID.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtDID.Name = "txtDID";
            this.txtDID.Size = new System.Drawing.Size(616, 26);
            this.txtDID.TabIndex = 1;
            this.txtDID.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDID_KeyDown);
            // 
            // lblDID
            // 
            this.lblDID.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.lblDID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblDID.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblDID.Location = new System.Drawing.Point(4, 160);
            this.lblDID.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblDID.Name = "lblDID";
            this.lblDID.Size = new System.Drawing.Size(86, 24);
            this.lblDID.TabIndex = 0;
            this.lblDID.Text = "唯一码:";
            this.lblDID.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.groupboxLabel);
            this.groupBox2.Controls.Add(this.btnExit);
            this.groupBox2.Controls.Add(this.btnQuery);
            this.groupBox2.Controls.Add(this.btnReprint);
            this.groupBox2.Controls.Add(this.btnExcel);
            this.groupBox2.Controls.Add(this.btnSave);
            this.groupBox2.Controls.Add(this.btnRefresh);
            this.groupBox2.Controls.Add(this.btnCancel);
            this.groupBox2.Location = new System.Drawing.Point(9, 265);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox2.Size = new System.Drawing.Size(724, 64);
            this.groupBox2.TabIndex = 4;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "按钮";
            // 
            // groupboxLabel
            // 
            this.groupboxLabel.Controls.Add(this.OptOld);
            this.groupboxLabel.Controls.Add(this.OptNew);
            this.groupboxLabel.ForeColor = System.Drawing.Color.Red;
            this.groupboxLabel.Location = new System.Drawing.Point(579, 13);
            this.groupboxLabel.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupboxLabel.Name = "groupboxLabel";
            this.groupboxLabel.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupboxLabel.Size = new System.Drawing.Size(130, 39);
            this.groupboxLabel.TabIndex = 25;
            this.groupboxLabel.TabStop = false;
            this.groupboxLabel.Text = "Label";
            // 
            // OptOld
            // 
            this.OptOld.AutoSize = true;
            this.OptOld.Location = new System.Drawing.Point(18, 18);
            this.OptOld.Name = "OptOld";
            this.OptOld.Size = new System.Drawing.Size(41, 16);
            this.OptOld.TabIndex = 186;
            this.OptOld.Text = "Old";
            this.OptOld.UseVisualStyleBackColor = true;
            // 
            // OptNew
            // 
            this.OptNew.AutoSize = true;
            this.OptNew.Checked = true;
            this.OptNew.Location = new System.Drawing.Point(75, 18);
            this.OptNew.Name = "OptNew";
            this.OptNew.Size = new System.Drawing.Size(41, 16);
            this.OptNew.TabIndex = 185;
            this.OptNew.TabStop = true;
            this.OptNew.Text = "New";
            this.OptNew.UseVisualStyleBackColor = true;
            // 
            // btnExit
            // 
            this.btnExit.AutoSize = true;
            this.btnExit.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.btnExit.Enabled = false;
            this.btnExit.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnExit.Font = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnExit.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.btnExit.Location = new System.Drawing.Point(432, 20);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(52, 32);
            this.btnExit.TabIndex = 24;
            this.btnExit.Text = "退出";
            this.btnExit.UseVisualStyleBackColor = false;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnQuery
            // 
            this.btnQuery.AutoSize = true;
            this.btnQuery.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.btnQuery.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnQuery.Font = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnQuery.ForeColor = System.Drawing.Color.Blue;
            this.btnQuery.Location = new System.Drawing.Point(13, 20);
            this.btnQuery.Name = "btnQuery";
            this.btnQuery.Size = new System.Drawing.Size(52, 32);
            this.btnQuery.TabIndex = 15;
            this.btnQuery.Text = "查询";
            this.btnQuery.UseVisualStyleBackColor = false;
            this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // btnReprint
            // 
            this.btnReprint.AutoSize = true;
            this.btnReprint.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.btnReprint.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnReprint.Font = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnReprint.ForeColor = System.Drawing.Color.Olive;
            this.btnReprint.Location = new System.Drawing.Point(364, 20);
            this.btnReprint.Name = "btnReprint";
            this.btnReprint.Size = new System.Drawing.Size(52, 32);
            this.btnReprint.TabIndex = 23;
            this.btnReprint.Text = "补印";
            this.btnReprint.UseVisualStyleBackColor = false;
            // 
            // btnExcel
            // 
            this.btnExcel.AutoSize = true;
            this.btnExcel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.btnExcel.Enabled = false;
            this.btnExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnExcel.Font = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnExcel.ForeColor = System.Drawing.Color.Purple;
            this.btnExcel.Location = new System.Drawing.Point(294, 20);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(52, 32);
            this.btnExcel.TabIndex = 22;
            this.btnExcel.Text = "导出";
            this.btnExcel.UseVisualStyleBackColor = false;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // btnSave
            // 
            this.btnSave.AutoSize = true;
            this.btnSave.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.btnSave.Enabled = false;
            this.btnSave.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSave.Font = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnSave.ForeColor = System.Drawing.Color.Green;
            this.btnSave.Location = new System.Drawing.Point(81, 20);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(52, 32);
            this.btnSave.TabIndex = 19;
            this.btnSave.Text = "保存";
            this.btnSave.UseVisualStyleBackColor = false;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnRefresh
            // 
            this.btnRefresh.AutoSize = true;
            this.btnRefresh.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.btnRefresh.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnRefresh.Font = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnRefresh.ForeColor = System.Drawing.Color.Navy;
            this.btnRefresh.Location = new System.Drawing.Point(223, 20);
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.Size = new System.Drawing.Size(52, 32);
            this.btnRefresh.TabIndex = 21;
            this.btnRefresh.Text = "刷新";
            this.btnRefresh.UseVisualStyleBackColor = false;
            this.btnRefresh.Click += new System.EventHandler(this.btnRefresh_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.AutoSize = true;
            this.btnCancel.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.btnCancel.Enabled = false;
            this.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnCancel.Font = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnCancel.ForeColor = System.Drawing.Color.Teal;
            this.btnCancel.Location = new System.Drawing.Point(148, 20);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(52, 32);
            this.btnCancel.TabIndex = 20;
            this.btnCancel.Text = "取消";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // DG_Result
            // 
            this.DG_Result.AllowUserToAddRows = false;
            this.DG_Result.AllowUserToDeleteRows = false;
            this.DG_Result.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.DG_Result.BackgroundColor = System.Drawing.SystemColors.Control;
            this.DG_Result.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DG_Result.Location = new System.Drawing.Point(9, 334);
            this.DG_Result.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.DG_Result.Name = "DG_Result";
            this.DG_Result.RowTemplate.Height = 27;
            this.DG_Result.Size = new System.Drawing.Size(724, 262);
            this.DG_Result.TabIndex = 17;
            // 
            // groupBox3
            // 
            this.groupBox3.BackColor = System.Drawing.Color.Linen;
            this.groupBox3.Controls.Add(this.lblMsg);
            this.groupBox3.Location = new System.Drawing.Point(10, 203);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox3.Size = new System.Drawing.Size(724, 57);
            this.groupBox3.TabIndex = 49;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Message";
            // 
            // lblMsg
            // 
            this.lblMsg.BackColor = System.Drawing.Color.White;
            this.lblMsg.Font = new System.Drawing.Font("宋体", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblMsg.ForeColor = System.Drawing.Color.Red;
            this.lblMsg.Location = new System.Drawing.Point(4, 18);
            this.lblMsg.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(704, 30);
            this.lblMsg.TabIndex = 4;
            this.lblMsg.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // frmDIDInteGration
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(752, 641);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.DG_Result);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "frmDIDInteGration";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmDIDInteGration";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmDIDInteGration_FormClosed);
            this.Load += new System.EventHandler(this.frmDIDInteGration_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupboxLabel.ResumeLayout(false);
            this.groupboxLabel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DG_Result)).EndInit();
            this.groupBox3.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtVendorCode;
        private System.Windows.Forms.Label lblVendorCode;
        private System.Windows.Forms.TextBox txtLotCode;
        private System.Windows.Forms.Label lblLotCode;
        private System.Windows.Forms.Label lblCompPN;
        private System.Windows.Forms.TextBox txtDateCode;
        private System.Windows.Forms.Label lblDateCode;
        private System.Windows.Forms.TextBox txtDID;
        private System.Windows.Forms.Label lblDID;
        private System.Windows.Forms.TextBox txtEndT;
        private System.Windows.Forms.TextBox txtBeginT;
        private System.Windows.Forms.DateTimePicker DTPEndDate;
        private System.Windows.Forms.Label lblEndT;
        private System.Windows.Forms.DateTimePicker DTPBeginDate;
        private System.Windows.Forms.Label lblBeginT;
        private System.Windows.Forms.Label lblLine;
        private System.Windows.Forms.ComboBox cboLine;
        private System.Windows.Forms.Label lblSide;
        private System.Windows.Forms.ComboBox cboSide;
        private System.Windows.Forms.Label lblFactory;
        private System.Windows.Forms.ComboBox cboFactory;
        private System.Windows.Forms.TextBox txtQty;
        private System.Windows.Forms.Label lblQty;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.Button btnQuery;
        private System.Windows.Forms.Button btnReprint;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnRefresh;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.DataGridView DG_Result;
        private System.Windows.Forms.GroupBox groupboxLabel;
        private System.Windows.Forms.RadioButton OptOld;
        private System.Windows.Forms.RadioButton OptNew;
        private System.Windows.Forms.TextBox txtCompPN;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label lblMsg;
    }
}