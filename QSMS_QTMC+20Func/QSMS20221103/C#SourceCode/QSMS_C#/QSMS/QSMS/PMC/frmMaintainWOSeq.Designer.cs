namespace QSMS.QSMS.PMC
{
    partial class frmMaintainWOSeq
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
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnReset = new System.Windows.Forms.Button();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnDelete = new System.Windows.Forms.Button();
            this.CboGroupID = new System.Windows.Forms.ComboBox();
            this.CboWo = new System.Windows.Forms.ComboBox();
            this.txtWO = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.rbtnGroup = new System.Windows.Forms.RadioButton();
            this.rbtnRelease = new System.Windows.Forms.RadioButton();
            this.btnQueryID = new System.Windows.Forms.Button();
            this.btnDELALL = new System.Windows.Forms.Button();
            this.btnDEL = new System.Windows.Forms.Button();
            this.btnADDALL = new System.Windows.Forms.Button();
            this.btnAdd = new System.Windows.Forms.Button();
            this.lstWO_SELECT = new System.Windows.Forms.ListBox();
            this.lstWO_LIST = new System.Windows.Forms.ListBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.btnQuery = new System.Windows.Forms.Button();
            this.txtWOQty = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtMBPN = new System.Windows.Forms.TextBox();
            this.CboLine = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.dtpEDate = new System.Windows.Forms.DateTimePicker();
            this.dtpSDate = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.Info;
            this.label2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(5, 14);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(80, 26);
            this.label2.TabIndex = 1;
            this.label2.Text = "BeginDate";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnReset);
            this.groupBox1.Controls.Add(this.btnUpdate);
            this.groupBox1.Controls.Add(this.btnSave);
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Controls.Add(this.btnDELALL);
            this.groupBox1.Controls.Add(this.btnDEL);
            this.groupBox1.Controls.Add(this.btnADDALL);
            this.groupBox1.Controls.Add(this.btnAdd);
            this.groupBox1.Controls.Add(this.lstWO_SELECT);
            this.groupBox1.Controls.Add(this.lstWO_LIST);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.btnQuery);
            this.groupBox1.Controls.Add(this.txtWOQty);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.txtMBPN);
            this.groupBox1.Controls.Add(this.CboLine);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.dtpEDate);
            this.groupBox1.Controls.Add(this.dtpSDate);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Location = new System.Drawing.Point(10, 4);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox1.Size = new System.Drawing.Size(707, 564);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            // 
            // btnReset
            // 
            this.btnReset.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnReset.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnReset.Location = new System.Drawing.Point(424, 398);
            this.btnReset.Margin = new System.Windows.Forms.Padding(2);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(98, 44);
            this.btnReset.TabIndex = 29;
            this.btnReset.Text = "Reset";
            this.btnReset.UseVisualStyleBackColor = false;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnUpdate.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnUpdate.Location = new System.Drawing.Point(424, 334);
            this.btnUpdate.Margin = new System.Windows.Forms.Padding(2);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(98, 44);
            this.btnUpdate.TabIndex = 28;
            this.btnUpdate.Text = "Add";
            this.btnUpdate.UseVisualStyleBackColor = false;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // btnSave
            // 
            this.btnSave.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnSave.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnSave.Location = new System.Drawing.Point(424, 275);
            this.btnSave.Margin = new System.Windows.Forms.Padding(2);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(98, 44);
            this.btnSave.TabIndex = 27;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = false;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnDelete);
            this.groupBox2.Controls.Add(this.CboGroupID);
            this.groupBox2.Controls.Add(this.CboWo);
            this.groupBox2.Controls.Add(this.txtWO);
            this.groupBox2.Controls.Add(this.label10);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.rbtnGroup);
            this.groupBox2.Controls.Add(this.rbtnRelease);
            this.groupBox2.Controls.Add(this.btnQueryID);
            this.groupBox2.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox2.Location = new System.Drawing.Point(424, 17);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox2.Size = new System.Drawing.Size(276, 241);
            this.groupBox2.TabIndex = 20;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Find By WO  Group ID";
            // 
            // btnDelete
            // 
            this.btnDelete.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnDelete.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnDelete.Location = new System.Drawing.Point(87, 189);
            this.btnDelete.Margin = new System.Windows.Forms.Padding(2);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(112, 40);
            this.btnDelete.TabIndex = 26;
            this.btnDelete.Text = "Delete";
            this.btnDelete.UseVisualStyleBackColor = false;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // CboGroupID
            // 
            this.CboGroupID.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CboGroupID.FormattingEnabled = true;
            this.CboGroupID.Location = new System.Drawing.Point(54, 71);
            this.CboGroupID.Margin = new System.Windows.Forms.Padding(2);
            this.CboGroupID.Name = "CboGroupID";
            this.CboGroupID.Size = new System.Drawing.Size(218, 24);
            this.CboGroupID.TabIndex = 25;
            this.CboGroupID.DropDown += new System.EventHandler(this.CboGroupID_DropDown);
            this.CboGroupID.SelectedIndexChanged += new System.EventHandler(this.CboGroupID_SelectedIndexChanged);
            // 
            // CboWo
            // 
            this.CboWo.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CboWo.FormattingEnabled = true;
            this.CboWo.Location = new System.Drawing.Point(54, 112);
            this.CboWo.Margin = new System.Windows.Forms.Padding(2);
            this.CboWo.Name = "CboWo";
            this.CboWo.Size = new System.Drawing.Size(218, 24);
            this.CboWo.TabIndex = 21;
            this.CboWo.SelectedIndexChanged += new System.EventHandler(this.CboWo_SelectedIndexChanged);
            // 
            // txtWO
            // 
            this.txtWO.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtWO.Location = new System.Drawing.Point(54, 153);
            this.txtWO.Margin = new System.Windows.Forms.Padding(2);
            this.txtWO.Name = "txtWO";
            this.txtWO.Size = new System.Drawing.Size(218, 26);
            this.txtWO.TabIndex = 21;
            // 
            // label10
            // 
            this.label10.BackColor = System.Drawing.SystemColors.Info;
            this.label10.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label10.ForeColor = System.Drawing.Color.Black;
            this.label10.Location = new System.Drawing.Point(5, 151);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(44, 26);
            this.label10.TabIndex = 21;
            this.label10.Text = "WO";
            this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label9
            // 
            this.label9.BackColor = System.Drawing.SystemColors.Info;
            this.label9.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label9.Location = new System.Drawing.Point(5, 110);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(44, 26);
            this.label9.TabIndex = 24;
            this.label9.Text = "WO";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label8
            // 
            this.label8.BackColor = System.Drawing.SystemColors.Info;
            this.label8.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label8.Location = new System.Drawing.Point(5, 67);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(44, 26);
            this.label8.TabIndex = 21;
            this.label8.Text = "ID";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // rbtnGroup
            // 
            this.rbtnGroup.BackColor = System.Drawing.SystemColors.Info;
            this.rbtnGroup.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rbtnGroup.Location = new System.Drawing.Point(87, 23);
            this.rbtnGroup.Margin = new System.Windows.Forms.Padding(2);
            this.rbtnGroup.Name = "rbtnGroup";
            this.rbtnGroup.Size = new System.Drawing.Size(69, 22);
            this.rbtnGroup.TabIndex = 23;
            this.rbtnGroup.TabStop = true;
            this.rbtnGroup.Text = "Group";
            this.rbtnGroup.UseVisualStyleBackColor = false;
            // 
            // rbtnRelease
            // 
            this.rbtnRelease.BackColor = System.Drawing.SystemColors.Info;
            this.rbtnRelease.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rbtnRelease.Location = new System.Drawing.Point(4, 23);
            this.rbtnRelease.Margin = new System.Windows.Forms.Padding(2);
            this.rbtnRelease.Name = "rbtnRelease";
            this.rbtnRelease.Size = new System.Drawing.Size(79, 22);
            this.rbtnRelease.TabIndex = 22;
            this.rbtnRelease.TabStop = true;
            this.rbtnRelease.Text = "Release";
            this.rbtnRelease.UseVisualStyleBackColor = false;
            // 
            // btnQueryID
            // 
            this.btnQueryID.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnQueryID.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnQueryID.Location = new System.Drawing.Point(160, 30);
            this.btnQueryID.Margin = new System.Windows.Forms.Padding(2);
            this.btnQueryID.Name = "btnQueryID";
            this.btnQueryID.Size = new System.Drawing.Size(112, 30);
            this.btnQueryID.TabIndex = 21;
            this.btnQueryID.Text = "Query ID";
            this.btnQueryID.UseVisualStyleBackColor = false;
            this.btnQueryID.Click += new System.EventHandler(this.btnQueryID_Click);
            // 
            // btnDELALL
            // 
            this.btnDELALL.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnDELALL.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnDELALL.Location = new System.Drawing.Point(191, 397);
            this.btnDELALL.Margin = new System.Windows.Forms.Padding(2);
            this.btnDELALL.Name = "btnDELALL";
            this.btnDELALL.Size = new System.Drawing.Size(36, 25);
            this.btnDELALL.TabIndex = 19;
            this.btnDELALL.Text = "<<";
            this.btnDELALL.UseVisualStyleBackColor = false;
            this.btnDELALL.Click += new System.EventHandler(this.btnDELALL_Click);
            // 
            // btnDEL
            // 
            this.btnDEL.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnDEL.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnDEL.Location = new System.Drawing.Point(189, 353);
            this.btnDEL.Margin = new System.Windows.Forms.Padding(2);
            this.btnDEL.Name = "btnDEL";
            this.btnDEL.Size = new System.Drawing.Size(36, 25);
            this.btnDEL.TabIndex = 18;
            this.btnDEL.Text = "<";
            this.btnDEL.UseVisualStyleBackColor = false;
            this.btnDEL.Click += new System.EventHandler(this.btnDEL_Click);
            // 
            // btnADDALL
            // 
            this.btnADDALL.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnADDALL.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnADDALL.Location = new System.Drawing.Point(189, 309);
            this.btnADDALL.Margin = new System.Windows.Forms.Padding(2);
            this.btnADDALL.Name = "btnADDALL";
            this.btnADDALL.Size = new System.Drawing.Size(36, 25);
            this.btnADDALL.TabIndex = 17;
            this.btnADDALL.Text = ">>";
            this.btnADDALL.UseVisualStyleBackColor = false;
            this.btnADDALL.Click += new System.EventHandler(this.btnADDALL_Click);
            // 
            // btnAdd
            // 
            this.btnAdd.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnAdd.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnAdd.Location = new System.Drawing.Point(189, 265);
            this.btnAdd.Margin = new System.Windows.Forms.Padding(2);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(36, 25);
            this.btnAdd.TabIndex = 16;
            this.btnAdd.Text = ">";
            this.btnAdd.UseVisualStyleBackColor = false;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // lstWO_SELECT
            // 
            this.lstWO_SELECT.Font = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lstWO_SELECT.FormattingEnabled = true;
            this.lstWO_SELECT.ItemHeight = 19;
            this.lstWO_SELECT.Location = new System.Drawing.Point(231, 155);
            this.lstWO_SELECT.Margin = new System.Windows.Forms.Padding(2);
            this.lstWO_SELECT.Name = "lstWO_SELECT";
            this.lstWO_SELECT.Size = new System.Drawing.Size(180, 384);
            this.lstWO_SELECT.TabIndex = 15;
            // 
            // lstWO_LIST
            // 
            this.lstWO_LIST.Font = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lstWO_LIST.FormattingEnabled = true;
            this.lstWO_LIST.ItemHeight = 19;
            this.lstWO_LIST.Location = new System.Drawing.Point(4, 155);
            this.lstWO_LIST.Margin = new System.Windows.Forms.Padding(2);
            this.lstWO_LIST.Name = "lstWO_LIST";
            this.lstWO_LIST.Size = new System.Drawing.Size(180, 384);
            this.lstWO_LIST.TabIndex = 14;
            this.lstWO_LIST.SelectedIndexChanged += new System.EventHandler(this.lstWO_LIST_SelectedIndexChanged);
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.SystemColors.Info;
            this.label7.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label7.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label7.Location = new System.Drawing.Point(232, 127);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(178, 26);
            this.label7.TabIndex = 13;
            this.label7.Text = "WO Selected Seq";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.SystemColors.Info;
            this.label6.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(64)))), ((int)(((byte)(0)))));
            this.label6.Location = new System.Drawing.Point(5, 127);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(201, 26);
            this.label6.TabIndex = 12;
            this.label6.Text = "Total WO---Without Group ID";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnQuery
            // 
            this.btnQuery.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnQuery.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnQuery.Location = new System.Drawing.Point(298, 85);
            this.btnQuery.Margin = new System.Windows.Forms.Padding(2);
            this.btnQuery.Name = "btnQuery";
            this.btnQuery.Size = new System.Drawing.Size(112, 30);
            this.btnQuery.TabIndex = 11;
            this.btnQuery.Text = "Query WO";
            this.btnQuery.UseVisualStyleBackColor = false;
            this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // txtWOQty
            // 
            this.txtWOQty.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtWOQty.Location = new System.Drawing.Point(91, 86);
            this.txtWOQty.Margin = new System.Windows.Forms.Padding(2);
            this.txtWOQty.Name = "txtWOQty";
            this.txtWOQty.Size = new System.Drawing.Size(204, 26);
            this.txtWOQty.TabIndex = 10;
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.SystemColors.Info;
            this.label5.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.ForeColor = System.Drawing.Color.Green;
            this.label5.Location = new System.Drawing.Point(5, 85);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(80, 26);
            this.label5.TabIndex = 9;
            this.label5.Text = "Qty";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtMBPN
            // 
            this.txtMBPN.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtMBPN.Location = new System.Drawing.Point(91, 51);
            this.txtMBPN.Margin = new System.Windows.Forms.Padding(2);
            this.txtMBPN.Name = "txtMBPN";
            this.txtMBPN.Size = new System.Drawing.Size(204, 26);
            this.txtMBPN.TabIndex = 8;
            // 
            // CboLine
            // 
            this.CboLine.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CboLine.FormattingEnabled = true;
            this.CboLine.Location = new System.Drawing.Point(347, 51);
            this.CboLine.Margin = new System.Windows.Forms.Padding(2);
            this.CboLine.Name = "CboLine";
            this.CboLine.Size = new System.Drawing.Size(63, 24);
            this.CboLine.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.SystemColors.Info;
            this.label4.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(298, 48);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(44, 26);
            this.label4.TabIndex = 6;
            this.label4.Text = "Line";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.SystemColors.Info;
            this.label3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.ForeColor = System.Drawing.Color.ForestGreen;
            this.label3.Location = new System.Drawing.Point(5, 50);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(80, 26);
            this.label3.TabIndex = 5;
            this.label3.Text = "MB PN";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // dtpEDate
            // 
            this.dtpEDate.CalendarFont = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dtpEDate.Checked = false;
            this.dtpEDate.Font = new System.Drawing.Font("宋体", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dtpEDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpEDate.Location = new System.Drawing.Point(298, 17);
            this.dtpEDate.Margin = new System.Windows.Forms.Padding(2);
            this.dtpEDate.Name = "dtpEDate";
            this.dtpEDate.Size = new System.Drawing.Size(113, 24);
            this.dtpEDate.TabIndex = 4;
            // 
            // dtpSDate
            // 
            this.dtpSDate.CalendarFont = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dtpSDate.Checked = false;
            this.dtpSDate.Font = new System.Drawing.Font("宋体", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dtpSDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpSDate.Location = new System.Drawing.Point(91, 17);
            this.dtpSDate.Margin = new System.Windows.Forms.Padding(2);
            this.dtpSDate.Name = "dtpSDate";
            this.dtpSDate.Size = new System.Drawing.Size(105, 24);
            this.dtpSDate.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(213, 14);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 26);
            this.label1.TabIndex = 2;
            this.label1.Text = "End Date";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // frmMaintainWOSeq
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(724, 578);
            this.Controls.Add(this.groupBox1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "frmMaintainWOSeq";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Maintain WO Seq [20201222]";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmMaintainWOSeq_FormClosed);
            this.Load += new System.EventHandler(this.frmMaintainWOSeq_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtMBPN;
        private System.Windows.Forms.ComboBox CboLine;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker dtpEDate;
        private System.Windows.Forms.DateTimePicker dtpSDate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.ComboBox CboGroupID;
        private System.Windows.Forms.ComboBox CboWo;
        private System.Windows.Forms.TextBox txtWO;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.RadioButton rbtnGroup;
        private System.Windows.Forms.RadioButton rbtnRelease;
        private System.Windows.Forms.Button btnQueryID;
        private System.Windows.Forms.Button btnDELALL;
        private System.Windows.Forms.Button btnDEL;
        private System.Windows.Forms.Button btnADDALL;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.ListBox lstWO_SELECT;
        private System.Windows.Forms.ListBox lstWO_LIST;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Button btnQuery;
        private System.Windows.Forms.TextBox txtWOQty;
        private System.Windows.Forms.Label label5;

    }
}