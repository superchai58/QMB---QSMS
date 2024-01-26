namespace QSMS.QSMS.MCC
{
    partial class FrmSingleSideBrdConfirm
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
            this.OptRelease = new System.Windows.Forms.RadioButton();
            this.optGroup = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.dtpSDate = new System.Windows.Forms.DateTimePicker();
            this.dtpEDate = new System.Windows.Forms.DateTimePicker();
            this.CboLine = new System.Windows.Forms.ComboBox();
            this.CmdQuery = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.CboGroupID = new System.Windows.Forms.ComboBox();
            this.CboNotChkBOM = new System.Windows.Forms.ComboBox();
            this.cboWO = new System.Windows.Forms.ComboBox();
            this.CboNotFinishedWO = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.CboSBWO = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.TxtWO = new System.Windows.Forms.TextBox();
            this.TxtMBPN = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.TxtCustomer = new System.Windows.Forms.TextBox();
            this.TxtGroup = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.TxtModel = new System.Windows.Forms.TextBox();
            this.TxtWOQty = new System.Windows.Forms.TextBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.CmdConfirm = new System.Windows.Forms.Button();
            this.CboBuildType = new System.Windows.Forms.ComboBox();
            this.label14 = new System.Windows.Forms.Label();
            this.Opt_Both = new System.Windows.Forms.RadioButton();
            this.Opt_Solder = new System.Windows.Forms.RadioButton();
            this.Opt_Comp = new System.Windows.Forms.RadioButton();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // OptRelease
            // 
            this.OptRelease.BackColor = System.Drawing.SystemColors.Info;
            this.OptRelease.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.OptRelease.Location = new System.Drawing.Point(19, 13);
            this.OptRelease.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.OptRelease.Name = "OptRelease";
            this.OptRelease.Size = new System.Drawing.Size(88, 21);
            this.OptRelease.TabIndex = 0;
            this.OptRelease.TabStop = true;
            this.OptRelease.Text = "Release";
            this.OptRelease.UseVisualStyleBackColor = false;
            // 
            // optGroup
            // 
            this.optGroup.BackColor = System.Drawing.SystemColors.Info;
            this.optGroup.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.optGroup.Location = new System.Drawing.Point(115, 13);
            this.optGroup.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.optGroup.Name = "optGroup";
            this.optGroup.Size = new System.Drawing.Size(77, 21);
            this.optGroup.TabIndex = 1;
            this.optGroup.TabStop = true;
            this.optGroup.Tag = "";
            this.optGroup.Text = "Group";
            this.optGroup.UseVisualStyleBackColor = false;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(18, 52);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(87, 24);
            this.label1.TabIndex = 2;
            this.label1.Text = "BeginDate";
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.Info;
            this.label2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(256, 49);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(88, 24);
            this.label2.TabIndex = 3;
            this.label2.Text = "End Date";
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.SystemColors.Info;
            this.label3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(499, 49);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(53, 24);
            this.label3.TabIndex = 4;
            this.label3.Text = "Line";
            // 
            // dtpSDate
            // 
            this.dtpSDate.Location = new System.Drawing.Point(109, 49);
            this.dtpSDate.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.dtpSDate.Name = "dtpSDate";
            this.dtpSDate.Size = new System.Drawing.Size(139, 26);
            this.dtpSDate.TabIndex = 5;
            // 
            // dtpEDate
            // 
            this.dtpEDate.Location = new System.Drawing.Point(352, 49);
            this.dtpEDate.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.dtpEDate.Name = "dtpEDate";
            this.dtpEDate.Size = new System.Drawing.Size(139, 26);
            this.dtpEDate.TabIndex = 6;
            // 
            // CboLine
            // 
            this.CboLine.FormattingEnabled = true;
            this.CboLine.Location = new System.Drawing.Point(557, 49);
            this.CboLine.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.CboLine.Name = "CboLine";
            this.CboLine.Size = new System.Drawing.Size(139, 24);
            this.CboLine.TabIndex = 7;
            // 
            // CmdQuery
            // 
            this.CmdQuery.Location = new System.Drawing.Point(718, 47);
            this.CmdQuery.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.CmdQuery.Name = "CmdQuery";
            this.CmdQuery.Size = new System.Drawing.Size(104, 27);
            this.CmdQuery.TabIndex = 8;
            this.CmdQuery.Text = "&Find";
            this.CmdQuery.UseVisualStyleBackColor = true;
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.SystemColors.Info;
            this.label4.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(18, 90);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(87, 24);
            this.label4.TabIndex = 9;
            this.label4.Text = "GroupID";
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.SystemColors.Info;
            this.label5.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.Location = new System.Drawing.Point(256, 93);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(115, 21);
            this.label5.TabIndex = 10;
            this.label5.Text = "No Chk BOM OK";
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.SystemColors.Info;
            this.label6.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.Location = new System.Drawing.Point(506, 90);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(48, 24);
            this.label6.TabIndex = 11;
            this.label6.Text = "OK WO";
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.SystemColors.Info;
            this.label7.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label7.Location = new System.Drawing.Point(18, 122);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(89, 24);
            this.label7.TabIndex = 12;
            this.label7.Text = "Un OK WO";
            // 
            // CboGroupID
            // 
            this.CboGroupID.FormattingEnabled = true;
            this.CboGroupID.Location = new System.Drawing.Point(109, 90);
            this.CboGroupID.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.CboGroupID.Name = "CboGroupID";
            this.CboGroupID.Size = new System.Drawing.Size(139, 24);
            this.CboGroupID.TabIndex = 13;
            // 
            // CboNotChkBOM
            // 
            this.CboNotChkBOM.FormattingEnabled = true;
            this.CboNotChkBOM.Location = new System.Drawing.Point(376, 90);
            this.CboNotChkBOM.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.CboNotChkBOM.Name = "CboNotChkBOM";
            this.CboNotChkBOM.Size = new System.Drawing.Size(122, 24);
            this.CboNotChkBOM.TabIndex = 14;
            // 
            // cboWO
            // 
            this.cboWO.FormattingEnabled = true;
            this.cboWO.Location = new System.Drawing.Point(559, 90);
            this.cboWO.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cboWO.Name = "cboWO";
            this.cboWO.Size = new System.Drawing.Size(137, 24);
            this.cboWO.TabIndex = 15;
            // 
            // CboNotFinishedWO
            // 
            this.CboNotFinishedWO.FormattingEnabled = true;
            this.CboNotFinishedWO.Location = new System.Drawing.Point(109, 122);
            this.CboNotFinishedWO.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.CboNotFinishedWO.Name = "CboNotFinishedWO";
            this.CboNotFinishedWO.Size = new System.Drawing.Size(139, 24);
            this.CboNotFinishedWO.TabIndex = 16;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.CboSBWO);
            this.groupBox1.Location = new System.Drawing.Point(704, 82);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox1.Size = new System.Drawing.Size(152, 53);
            this.groupBox1.TabIndex = 17;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Small Board WO";
            // 
            // CboSBWO
            // 
            this.CboSBWO.FormattingEnabled = true;
            this.CboSBWO.Location = new System.Drawing.Point(9, 18);
            this.CboSBWO.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.CboSBWO.Name = "CboSBWO";
            this.CboSBWO.Size = new System.Drawing.Size(135, 24);
            this.CboSBWO.TabIndex = 0;
            // 
            // label8
            // 
            this.label8.BackColor = System.Drawing.SystemColors.Info;
            this.label8.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label8.Location = new System.Drawing.Point(253, 123);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(44, 23);
            this.label8.TabIndex = 18;
            this.label8.Text = "WO";
            // 
            // label9
            // 
            this.label9.BackColor = System.Drawing.SystemColors.Info;
            this.label9.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label9.Location = new System.Drawing.Point(445, 121);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(56, 23);
            this.label9.TabIndex = 19;
            this.label9.Text = "MB PN";
            // 
            // TxtWO
            // 
            this.TxtWO.Location = new System.Drawing.Point(298, 120);
            this.TxtWO.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.TxtWO.Name = "TxtWO";
            this.TxtWO.Size = new System.Drawing.Size(139, 26);
            this.TxtWO.TabIndex = 20;
            // 
            // TxtMBPN
            // 
            this.TxtMBPN.Location = new System.Drawing.Point(509, 118);
            this.TxtMBPN.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.TxtMBPN.Name = "TxtMBPN";
            this.TxtMBPN.Size = new System.Drawing.Size(139, 26);
            this.TxtMBPN.TabIndex = 21;
            // 
            // label10
            // 
            this.label10.BackColor = System.Drawing.SystemColors.Info;
            this.label10.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label10.Location = new System.Drawing.Point(18, 161);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(78, 23);
            this.label10.TabIndex = 22;
            this.label10.Text = "Customer";
            // 
            // label11
            // 
            this.label11.BackColor = System.Drawing.SystemColors.Info;
            this.label11.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label11.Location = new System.Drawing.Point(209, 161);
            this.label11.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(97, 23);
            this.label11.TabIndex = 23;
            this.label11.Text = "Group(M/S)";
            // 
            // TxtCustomer
            // 
            this.TxtCustomer.Location = new System.Drawing.Point(91, 158);
            this.TxtCustomer.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.TxtCustomer.Name = "TxtCustomer";
            this.TxtCustomer.Size = new System.Drawing.Size(113, 26);
            this.TxtCustomer.TabIndex = 24;
            // 
            // TxtGroup
            // 
            this.TxtGroup.Location = new System.Drawing.Point(309, 161);
            this.TxtGroup.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.TxtGroup.Name = "TxtGroup";
            this.TxtGroup.Size = new System.Drawing.Size(95, 26);
            this.TxtGroup.TabIndex = 25;
            // 
            // label12
            // 
            this.label12.BackColor = System.Drawing.SystemColors.Info;
            this.label12.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label12.Location = new System.Drawing.Point(412, 161);
            this.label12.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(56, 23);
            this.label12.TabIndex = 26;
            this.label12.Text = "Model";
            // 
            // label13
            // 
            this.label13.BackColor = System.Drawing.SystemColors.Info;
            this.label13.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label13.Location = new System.Drawing.Point(599, 161);
            this.label13.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(35, 23);
            this.label13.TabIndex = 27;
            this.label13.Text = "Qty";
            // 
            // TxtModel
            // 
            this.TxtModel.Location = new System.Drawing.Point(466, 161);
            this.TxtModel.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.TxtModel.Name = "TxtModel";
            this.TxtModel.Size = new System.Drawing.Size(132, 26);
            this.TxtModel.TabIndex = 28;
            // 
            // TxtWOQty
            // 
            this.TxtWOQty.Location = new System.Drawing.Point(640, 161);
            this.TxtWOQty.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.TxtWOQty.Name = "TxtWOQty";
            this.TxtWOQty.Size = new System.Drawing.Size(76, 26);
            this.TxtWOQty.TabIndex = 29;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.CmdConfirm);
            this.groupBox2.Controls.Add(this.CboBuildType);
            this.groupBox2.Controls.Add(this.label14);
            this.groupBox2.Controls.Add(this.Opt_Both);
            this.groupBox2.Controls.Add(this.Opt_Solder);
            this.groupBox2.Controls.Add(this.Opt_Comp);
            this.groupBox2.Location = new System.Drawing.Point(16, 195);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox2.Size = new System.Drawing.Size(700, 133);
            this.groupBox2.TabIndex = 30;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Confirm Single Side Brd";
            // 
            // CmdConfirm
            // 
            this.CmdConfirm.Location = new System.Drawing.Point(233, 73);
            this.CmdConfirm.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.CmdConfirm.Name = "CmdConfirm";
            this.CmdConfirm.Size = new System.Drawing.Size(79, 31);
            this.CmdConfirm.TabIndex = 5;
            this.CmdConfirm.Text = "&Save";
            this.CmdConfirm.UseVisualStyleBackColor = true;
            // 
            // CboBuildType
            // 
            this.CboBuildType.FormattingEnabled = true;
            this.CboBuildType.Location = new System.Drawing.Point(102, 73);
            this.CboBuildType.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.CboBuildType.Name = "CboBuildType";
            this.CboBuildType.Size = new System.Drawing.Size(123, 24);
            this.CboBuildType.TabIndex = 4;
            // 
            // label14
            // 
            this.label14.BackColor = System.Drawing.SystemColors.Info;
            this.label14.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label14.Location = new System.Drawing.Point(18, 73);
            this.label14.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(79, 27);
            this.label14.TabIndex = 3;
            this.label14.Text = "BuildType";
            // 
            // Opt_Both
            // 
            this.Opt_Both.BackColor = System.Drawing.SystemColors.Info;
            this.Opt_Both.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Opt_Both.Location = new System.Drawing.Point(299, 31);
            this.Opt_Both.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Opt_Both.Name = "Opt_Both";
            this.Opt_Both.Size = new System.Drawing.Size(103, 21);
            this.Opt_Both.TabIndex = 2;
            this.Opt_Both.TabStop = true;
            this.Opt_Both.Text = "Both Side";
            this.Opt_Both.UseVisualStyleBackColor = false;
            // 
            // Opt_Solder
            // 
            this.Opt_Solder.BackColor = System.Drawing.SystemColors.Info;
            this.Opt_Solder.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Opt_Solder.Location = new System.Drawing.Point(172, 31);
            this.Opt_Solder.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Opt_Solder.Name = "Opt_Solder";
            this.Opt_Solder.Size = new System.Drawing.Size(119, 21);
            this.Opt_Solder.TabIndex = 1;
            this.Opt_Solder.TabStop = true;
            this.Opt_Solder.Text = "Solder Side";
            this.Opt_Solder.UseVisualStyleBackColor = false;
            // 
            // Opt_Comp
            // 
            this.Opt_Comp.BackColor = System.Drawing.SystemColors.Info;
            this.Opt_Comp.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Opt_Comp.Location = new System.Drawing.Point(21, 31);
            this.Opt_Comp.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Opt_Comp.Name = "Opt_Comp";
            this.Opt_Comp.Size = new System.Drawing.Size(143, 21);
            this.Opt_Comp.TabIndex = 0;
            this.Opt_Comp.TabStop = true;
            this.Opt_Comp.Text = "Component Side";
            this.Opt_Comp.UseVisualStyleBackColor = false;
            // 
            // FrmSingleSideBrdConfirm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(880, 349);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.TxtWOQty);
            this.Controls.Add(this.TxtModel);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.TxtGroup);
            this.Controls.Add(this.TxtCustomer);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.TxtMBPN);
            this.Controls.Add(this.TxtWO);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.CboNotFinishedWO);
            this.Controls.Add(this.cboWO);
            this.Controls.Add(this.CboNotChkBOM);
            this.Controls.Add(this.CboGroupID);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.CmdQuery);
            this.Controls.Add(this.CboLine);
            this.Controls.Add(this.dtpEDate);
            this.Controls.Add(this.dtpSDate);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.optGroup);
            this.Controls.Add(this.OptRelease);
            this.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "FrmSingleSideBrdConfirm";
            this.Text = "SingleSideBrdConfirm[2021/09/26]";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FrmSingleSideBrdConfirm_FormClosed);
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RadioButton OptRelease;
        private System.Windows.Forms.RadioButton optGroup;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker dtpSDate;
        private System.Windows.Forms.DateTimePicker dtpEDate;
        private System.Windows.Forms.ComboBox CboLine;
        private System.Windows.Forms.Button CmdQuery;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox CboGroupID;
        private System.Windows.Forms.ComboBox CboNotChkBOM;
        private System.Windows.Forms.ComboBox cboWO;
        private System.Windows.Forms.ComboBox CboNotFinishedWO;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox CboSBWO;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox TxtWO;
        private System.Windows.Forms.TextBox TxtMBPN;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox TxtCustomer;
        private System.Windows.Forms.TextBox TxtGroup;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox TxtModel;
        private System.Windows.Forms.TextBox TxtWOQty;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton Opt_Comp;
        private System.Windows.Forms.Button CmdConfirm;
        private System.Windows.Forms.ComboBox CboBuildType;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.RadioButton Opt_Both;
        private System.Windows.Forms.RadioButton Opt_Solder;
    }
}