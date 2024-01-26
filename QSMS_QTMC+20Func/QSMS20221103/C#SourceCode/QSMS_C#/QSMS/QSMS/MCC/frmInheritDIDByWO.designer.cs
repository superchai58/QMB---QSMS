namespace QSMS.QSMS.MCC
{
    partial class frmInheritDIDByWO
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmInheritDIDByWO));
            this.label5 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.OptSide = new System.Windows.Forms.RadioButton();
            this.CboInheritingWO = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.CboInheritWO = new System.Windows.Forms.ComboBox();
            this.OptMachine = new System.Windows.Forms.RadioButton();
            this.chkIncludeXL = new System.Windows.Forms.CheckBox();
            this.CboNotFinishedWO = new System.Windows.Forms.ComboBox();
            this.CboWO = new System.Windows.Forms.ComboBox();
            this.CboNotChkBOM = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.btsave = new System.Windows.Forms.Button();
            this.CboGroupID = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.CboLine = new System.Windows.Forms.ComboBox();
            this.btFind = new System.Windows.Forms.Button();
            this.dptEnd = new System.Windows.Forms.DateTimePicker();
            this.dptBegin = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtWO = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.txtGroup = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.txtWOQty = new System.Windows.Forms.TextBox();
            this.txtMBPN = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.lblmsg = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.SystemColors.Info;
            this.label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label5.Font = new System.Drawing.Font("宋体", 10.5F);
            this.label5.Location = new System.Drawing.Point(6, 49);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(80, 23);
            this.label5.TabIndex = 5;
            this.label5.Text = "BeginDate";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.OptSide);
            this.groupBox1.Controls.Add(this.CboInheritingWO);
            this.groupBox1.Controls.Add(this.label10);
            this.groupBox1.Controls.Add(this.label9);
            this.groupBox1.Controls.Add(this.CboInheritWO);
            this.groupBox1.Controls.Add(this.OptMachine);
            this.groupBox1.Controls.Add(this.chkIncludeXL);
            this.groupBox1.Controls.Add(this.CboNotFinishedWO);
            this.groupBox1.Controls.Add(this.CboWO);
            this.groupBox1.Controls.Add(this.CboNotChkBOM);
            this.groupBox1.Controls.Add(this.label8);
            this.groupBox1.Controls.Add(this.label7);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.btsave);
            this.groupBox1.Controls.Add(this.CboGroupID);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.CboLine);
            this.groupBox1.Controls.Add(this.btFind);
            this.groupBox1.Controls.Add(this.dptEnd);
            this.groupBox1.Controls.Add(this.dptBegin);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(945, 114);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            // 
            // OptSide
            // 
            this.OptSide.AutoSize = true;
            this.OptSide.BackColor = System.Drawing.SystemColors.Info;
            this.OptSide.Font = new System.Drawing.Font("宋体", 10.5F);
            this.OptSide.Location = new System.Drawing.Point(770, 48);
            this.OptSide.Name = "OptSide";
            this.OptSide.Size = new System.Drawing.Size(53, 18);
            this.OptSide.TabIndex = 184;
            this.OptSide.TabStop = true;
            this.OptSide.Text = "Side";
            this.OptSide.UseVisualStyleBackColor = false;
            // 
            // CboInheritingWO
            // 
            this.CboInheritingWO.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CboInheritingWO.FormattingEnabled = true;
            this.CboInheritingWO.Location = new System.Drawing.Point(593, 82);
            this.CboInheritingWO.Name = "CboInheritingWO";
            this.CboInheritingWO.Size = new System.Drawing.Size(155, 22);
            this.CboInheritingWO.TabIndex = 180;
            this.CboInheritingWO.SelectedIndexChanged += new System.EventHandler(this.CboInheritingWO_SelectedIndexChanged);
            // 
            // label10
            // 
            this.label10.BackColor = System.Drawing.SystemColors.Info;
            this.label10.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label10.Font = new System.Drawing.Font("宋体", 10.5F);
            this.label10.Location = new System.Drawing.Point(473, 83);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(114, 23);
            this.label10.TabIndex = 179;
            this.label10.Text = "Inheriting WO";
            this.label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label9
            // 
            this.label9.BackColor = System.Drawing.SystemColors.Info;
            this.label9.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label9.Font = new System.Drawing.Font("宋体", 10.5F);
            this.label9.Location = new System.Drawing.Point(192, 82);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(114, 23);
            this.label9.TabIndex = 178;
            this.label9.Text = "Inherit from WO";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // CboInheritWO
            // 
            this.CboInheritWO.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CboInheritWO.FormattingEnabled = true;
            this.CboInheritWO.Location = new System.Drawing.Point(312, 83);
            this.CboInheritWO.Name = "CboInheritWO";
            this.CboInheritWO.Size = new System.Drawing.Size(155, 22);
            this.CboInheritWO.TabIndex = 181;
            this.CboInheritWO.SelectedIndexChanged += new System.EventHandler(this.CboInheritWO_SelectedIndexChanged);
            // 
            // OptMachine
            // 
            this.OptMachine.AutoSize = true;
            this.OptMachine.BackColor = System.Drawing.SystemColors.Info;
            this.OptMachine.Font = new System.Drawing.Font("宋体", 10.5F);
            this.OptMachine.Location = new System.Drawing.Point(856, 48);
            this.OptMachine.Name = "OptMachine";
            this.OptMachine.Size = new System.Drawing.Size(74, 18);
            this.OptMachine.TabIndex = 183;
            this.OptMachine.TabStop = true;
            this.OptMachine.Text = "Machine";
            this.OptMachine.UseVisualStyleBackColor = false;
            // 
            // chkIncludeXL
            // 
            this.chkIncludeXL.AutoSize = true;
            this.chkIncludeXL.BackColor = System.Drawing.SystemColors.Info;
            this.chkIncludeXL.Font = new System.Drawing.Font("宋体", 10.5F);
            this.chkIncludeXL.Location = new System.Drawing.Point(770, 21);
            this.chkIncludeXL.Name = "chkIncludeXL";
            this.chkIncludeXL.Size = new System.Drawing.Size(145, 18);
            this.chkIncludeXL.TabIndex = 182;
            this.chkIncludeXL.Text = "Include XL CompPN";
            this.chkIncludeXL.UseVisualStyleBackColor = false;
            // 
            // CboNotFinishedWO
            // 
            this.CboNotFinishedWO.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CboNotFinishedWO.FormattingEnabled = true;
            this.CboNotFinishedWO.Location = new System.Drawing.Point(593, 50);
            this.CboNotFinishedWO.Name = "CboNotFinishedWO";
            this.CboNotFinishedWO.Size = new System.Drawing.Size(155, 22);
            this.CboNotFinishedWO.TabIndex = 177;
            this.CboNotFinishedWO.SelectedIndexChanged += new System.EventHandler(this.CboNotFinishedWO_SelectedIndexChanged);
            // 
            // CboWO
            // 
            this.CboWO.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CboWO.FormattingEnabled = true;
            this.CboWO.Location = new System.Drawing.Point(312, 51);
            this.CboWO.Name = "CboWO";
            this.CboWO.Size = new System.Drawing.Size(155, 22);
            this.CboWO.TabIndex = 176;
            this.CboWO.SelectedIndexChanged += new System.EventHandler(this.CboWO_SelectedIndexChanged);
            // 
            // CboNotChkBOM
            // 
            this.CboNotChkBOM.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CboNotChkBOM.FormattingEnabled = true;
            this.CboNotChkBOM.Location = new System.Drawing.Point(593, 17);
            this.CboNotChkBOM.Name = "CboNotChkBOM";
            this.CboNotChkBOM.Size = new System.Drawing.Size(155, 22);
            this.CboNotChkBOM.TabIndex = 175;
            this.CboNotChkBOM.SelectedIndexChanged += new System.EventHandler(this.CboNotChkBOM_SelectedIndexChanged);
            // 
            // label8
            // 
            this.label8.BackColor = System.Drawing.SystemColors.Info;
            this.label8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label8.Font = new System.Drawing.Font("宋体", 10.5F);
            this.label8.Location = new System.Drawing.Point(473, 17);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(114, 23);
            this.label8.TabIndex = 174;
            this.label8.Text = "ChkBomFailWO";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.SystemColors.Info;
            this.label7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label7.Font = new System.Drawing.Font("宋体", 10.5F);
            this.label7.Location = new System.Drawing.Point(192, 50);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(114, 23);
            this.label7.TabIndex = 173;
            this.label7.Text = "Finished WO";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.SystemColors.Info;
            this.label6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label6.Font = new System.Drawing.Font("宋体", 10.5F);
            this.label6.Location = new System.Drawing.Point(473, 51);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(114, 23);
            this.label6.TabIndex = 172;
            this.label6.Text = "Not Finished WO";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btsave
            // 
            this.btsave.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btsave.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btsave.FlatAppearance.BorderSize = 0;
            this.btsave.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btsave.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btsave.ForeColor = System.Drawing.Color.Black;
            this.btsave.Image = ((System.Drawing.Image)(resources.GetObject("btsave.Image")));
            this.btsave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btsave.Location = new System.Drawing.Point(854, 73);
            this.btsave.Name = "btsave";
            this.btsave.Size = new System.Drawing.Size(67, 30);
            this.btsave.TabIndex = 170;
            this.btsave.Text = "承接";
            this.btsave.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btsave.UseVisualStyleBackColor = false;
            this.btsave.Click += new System.EventHandler(this.btsave_Click);
            // 
            // CboGroupID
            // 
            this.CboGroupID.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CboGroupID.FormattingEnabled = true;
            this.CboGroupID.Location = new System.Drawing.Point(312, 17);
            this.CboGroupID.Name = "CboGroupID";
            this.CboGroupID.Size = new System.Drawing.Size(155, 22);
            this.CboGroupID.TabIndex = 169;
            this.CboGroupID.SelectedIndexChanged += new System.EventHandler(this.CboGroupID_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.SystemColors.Info;
            this.label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label3.Font = new System.Drawing.Font("宋体", 10.5F);
            this.label3.Location = new System.Drawing.Point(192, 17);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(114, 23);
            this.label3.TabIndex = 168;
            this.label3.Text = "GroupID";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.Info;
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label2.Font = new System.Drawing.Font("宋体", 10.5F);
            this.label2.Location = new System.Drawing.Point(6, 17);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(80, 23);
            this.label2.TabIndex = 167;
            this.label2.Text = "Line";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // CboLine
            // 
            this.CboLine.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CboLine.FormattingEnabled = true;
            this.CboLine.Location = new System.Drawing.Point(92, 17);
            this.CboLine.Name = "CboLine";
            this.CboLine.Size = new System.Drawing.Size(94, 22);
            this.CboLine.TabIndex = 166;
            // 
            // btFind
            // 
            this.btFind.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btFind.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btFind.FlatAppearance.BorderSize = 0;
            this.btFind.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btFind.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btFind.ForeColor = System.Drawing.Color.Black;
            this.btFind.Image = ((System.Drawing.Image)(resources.GetObject("btFind.Image")));
            this.btFind.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btFind.Location = new System.Drawing.Point(770, 72);
            this.btFind.Name = "btFind";
            this.btFind.Size = new System.Drawing.Size(67, 31);
            this.btFind.TabIndex = 165;
            this.btFind.Text = "查找";
            this.btFind.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.btFind.UseVisualStyleBackColor = false;
            this.btFind.Click += new System.EventHandler(this.btFind_Click);
            // 
            // dptEnd
            // 
            this.dptEnd.CalendarFont = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dptEnd.CustomFormat = "yyyy-MM-dd";
            this.dptEnd.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dptEnd.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dptEnd.Location = new System.Drawing.Point(92, 80);
            this.dptEnd.Name = "dptEnd";
            this.dptEnd.ShowUpDown = true;
            this.dptEnd.Size = new System.Drawing.Size(94, 23);
            this.dptEnd.TabIndex = 164;
            // 
            // dptBegin
            // 
            this.dptBegin.CustomFormat = "yyyy-MM-dd";
            this.dptBegin.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dptBegin.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dptBegin.Location = new System.Drawing.Point(92, 48);
            this.dptBegin.Name = "dptBegin";
            this.dptBegin.ShowUpDown = true;
            this.dptBegin.Size = new System.Drawing.Size(94, 23);
            this.dptBegin.TabIndex = 162;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label1.Font = new System.Drawing.Font("宋体", 10.5F);
            this.label1.Location = new System.Drawing.Point(6, 81);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 23);
            this.label1.TabIndex = 163;
            this.label1.Text = "EndData";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.SystemColors.Info;
            this.label4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label4.Font = new System.Drawing.Font("宋体", 10.5F);
            this.label4.Location = new System.Drawing.Point(6, 17);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(80, 26);
            this.label4.TabIndex = 185;
            this.label4.Text = "WO";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtWO
            // 
            this.txtWO.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtWO.Location = new System.Drawing.Point(96, 17);
            this.txtWO.Name = "txtWO";
            this.txtWO.ReadOnly = true;
            this.txtWO.Size = new System.Drawing.Size(126, 26);
            this.txtWO.TabIndex = 189;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.txtGroup);
            this.groupBox2.Controls.Add(this.label14);
            this.groupBox2.Controls.Add(this.txtWOQty);
            this.groupBox2.Controls.Add(this.txtMBPN);
            this.groupBox2.Controls.Add(this.label15);
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.txtWO);
            this.groupBox2.Location = new System.Drawing.Point(12, 149);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(945, 50);
            this.groupBox2.TabIndex = 191;
            this.groupBox2.TabStop = false;
            // 
            // txtGroup
            // 
            this.txtGroup.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtGroup.Location = new System.Drawing.Point(791, 15);
            this.txtGroup.Name = "txtGroup";
            this.txtGroup.ReadOnly = true;
            this.txtGroup.Size = new System.Drawing.Size(143, 26);
            this.txtGroup.TabIndex = 196;
            // 
            // label14
            // 
            this.label14.BackColor = System.Drawing.SystemColors.Info;
            this.label14.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label14.Font = new System.Drawing.Font("宋体", 10.5F);
            this.label14.Location = new System.Drawing.Point(701, 15);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(80, 26);
            this.label14.TabIndex = 193;
            this.label14.Text = "Group";
            this.label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtWOQty
            // 
            this.txtWOQty.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtWOQty.Location = new System.Drawing.Point(548, 16);
            this.txtWOQty.Name = "txtWOQty";
            this.txtWOQty.ReadOnly = true;
            this.txtWOQty.Size = new System.Drawing.Size(143, 26);
            this.txtWOQty.TabIndex = 195;
            // 
            // txtMBPN
            // 
            this.txtMBPN.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtMBPN.Location = new System.Drawing.Point(322, 16);
            this.txtMBPN.Name = "txtMBPN";
            this.txtMBPN.ReadOnly = true;
            this.txtMBPN.Size = new System.Drawing.Size(126, 26);
            this.txtMBPN.TabIndex = 197;
            // 
            // label15
            // 
            this.label15.BackColor = System.Drawing.SystemColors.Info;
            this.label15.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label15.Font = new System.Drawing.Font("宋体", 10.5F);
            this.label15.Location = new System.Drawing.Point(458, 15);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(80, 26);
            this.label15.TabIndex = 194;
            this.label15.Text = "Qty";
            this.label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label11
            // 
            this.label11.BackColor = System.Drawing.SystemColors.Info;
            this.label11.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label11.Font = new System.Drawing.Font("宋体", 10.5F);
            this.label11.Location = new System.Drawing.Point(232, 16);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(80, 26);
            this.label11.TabIndex = 191;
            this.label11.Text = "MBPN";
            this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblmsg
            // 
            this.lblmsg.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.lblmsg.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblmsg.Location = new System.Drawing.Point(12, 215);
            this.lblmsg.Name = "lblmsg";
            this.lblmsg.Size = new System.Drawing.Size(945, 36);
            this.lblmsg.TabIndex = 192;
            this.lblmsg.Text = "消息:";
            this.lblmsg.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // frmInheritDIDByWO
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(997, 290);
            this.Controls.Add(this.lblmsg);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.Name = "frmInheritDIDByWO";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "InheritDIDByWO";
            this.Load += new System.EventHandler(this.frmInheritDIDByWO_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dptEnd;
        private System.Windows.Forms.DateTimePicker dptBegin;
        private System.Windows.Forms.Button btFind;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox CboLine;
        private System.Windows.Forms.ComboBox CboGroupID;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btsave;
        private System.Windows.Forms.ComboBox CboInheritWO;
        private System.Windows.Forms.ComboBox CboInheritingWO;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox CboNotFinishedWO;
        private System.Windows.Forms.ComboBox CboWO;
        private System.Windows.Forms.ComboBox CboNotChkBOM;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.RadioButton OptSide;
        private System.Windows.Forms.RadioButton OptMachine;
        private System.Windows.Forms.CheckBox chkIncludeXL;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtWO;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox txtGroup;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox txtWOQty;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox txtMBPN;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label lblmsg;
    }
}