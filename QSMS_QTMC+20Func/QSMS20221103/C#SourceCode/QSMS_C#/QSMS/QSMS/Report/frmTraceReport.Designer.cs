namespace QSMS.QSMS.Report
{
    partial class frmTraceReport
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
            this.DataGridSN = new System.Windows.Forms.DataGridView();
            this.TxtSN = new System.Windows.Forms.TextBox();
            this.lblSNWO = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.OptBatch = new System.Windows.Forms.RadioButton();
            this.OptSN = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.CbbDataType = new System.Windows.Forms.ComboBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.opttxt = new System.Windows.Forms.RadioButton();
            this.optExcel = new System.Windows.Forms.RadioButton();
            this.cmdGetData = new System.Windows.Forms.Button();
            this.labelInfor = new System.Windows.Forms.Label();
            this.FraCompN = new System.Windows.Forms.GroupBox();
            this.lblInfor = new System.Windows.Forms.Label();
            this.DTPEndDate = new System.Windows.Forms.DateTimePicker();
            this.DTPBeginDate = new System.Windows.Forms.DateTimePicker();
            this.DTPBeginTime = new System.Windows.Forms.TextBox();
            this.DTPEndTime = new System.Windows.Forms.TextBox();
            this.txtVendorCode = new System.Windows.Forms.TextBox();
            this.txtLotCode = new System.Windows.Forms.TextBox();
            this.txtDateCode = new System.Windows.Forms.TextBox();
            this.txtModel = new System.Windows.Forms.TextBox();
            this.TxtCompPN = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.dtData = new System.Windows.Forms.DataGridView();
            this.inputSN = new System.Windows.Forms.Button();
            this.CMDChosefile = new System.Windows.Forms.Button();
            this.Txtpath = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label12 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DataGridSN)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.FraCompN.SuspendLayout();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dtData)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.DataGridSN);
            this.groupBox1.Controls.Add(this.TxtSN);
            this.groupBox1.Controls.Add(this.lblSNWO);
            this.groupBox1.Controls.Add(this.groupBox2);
            this.groupBox1.Location = new System.Drawing.Point(5, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(255, 542);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "FraSN";
            // 
            // DataGridSN
            // 
            this.DataGridSN.BackgroundColor = System.Drawing.Color.White;
            this.DataGridSN.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DataGridSN.Location = new System.Drawing.Point(7, 118);
            this.DataGridSN.Name = "DataGridSN";
            this.DataGridSN.RowTemplate.Height = 23;
            this.DataGridSN.Size = new System.Drawing.Size(242, 418);
            this.DataGridSN.TabIndex = 1;
            this.DataGridSN.MouseClick += new System.Windows.Forms.MouseEventHandler(this.DataGridSN_MouseClick);
            // 
            // TxtSN
            // 
            this.TxtSN.Location = new System.Drawing.Point(7, 91);
            this.TxtSN.Name = "TxtSN";
            this.TxtSN.Size = new System.Drawing.Size(240, 21);
            this.TxtSN.TabIndex = 1;
            // 
            // lblSNWO
            // 
            this.lblSNWO.AutoSize = true;
            this.lblSNWO.BackColor = System.Drawing.SystemColors.Control;
            this.lblSNWO.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblSNWO.Location = new System.Drawing.Point(7, 69);
            this.lblSNWO.Name = "lblSNWO";
            this.lblSNWO.Size = new System.Drawing.Size(240, 19);
            this.lblSNWO.TabIndex = 1;
            this.lblSNWO.Text = "SN/DID/WO:           ";
            // 
            // groupBox2
            // 
            this.groupBox2.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox2.Controls.Add(this.OptBatch);
            this.groupBox2.Controls.Add(this.OptSN);
            this.groupBox2.Location = new System.Drawing.Point(7, 20);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(242, 46);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            // 
            // OptBatch
            // 
            this.OptBatch.AutoSize = true;
            this.OptBatch.Location = new System.Drawing.Point(125, 20);
            this.OptBatch.Name = "OptBatch";
            this.OptBatch.Size = new System.Drawing.Size(71, 16);
            this.OptBatch.TabIndex = 2;
            this.OptBatch.Text = "By Batch";
            this.OptBatch.UseVisualStyleBackColor = true;
            // 
            // OptSN
            // 
            this.OptSN.AutoSize = true;
            this.OptSN.Checked = true;
            this.OptSN.Location = new System.Drawing.Point(23, 20);
            this.OptSN.Name = "OptSN";
            this.OptSN.Size = new System.Drawing.Size(59, 16);
            this.OptSN.TabIndex = 1;
            this.OptSN.TabStop = true;
            this.OptSN.Text = "By One";
            this.OptSN.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(276, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(108, 19);
            this.label1.TabIndex = 1;
            this.label1.Text = "DataType:";
            // 
            // CbbDataType
            // 
            this.CbbDataType.FormattingEnabled = true;
            this.CbbDataType.Location = new System.Drawing.Point(380, 22);
            this.CbbDataType.Name = "CbbDataType";
            this.CbbDataType.Size = new System.Drawing.Size(131, 20);
            this.CbbDataType.TabIndex = 2;
            this.CbbDataType.SelectedIndexChanged += new System.EventHandler(this.CbbDataType_SelectedIndexChanged);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.opttxt);
            this.groupBox3.Controls.Add(this.optExcel);
            this.groupBox3.Location = new System.Drawing.Point(526, 12);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(175, 41);
            this.groupBox3.TabIndex = 3;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "DataFormat";
            // 
            // opttxt
            // 
            this.opttxt.AutoSize = true;
            this.opttxt.Location = new System.Drawing.Point(110, 19);
            this.opttxt.Name = "opttxt";
            this.opttxt.Size = new System.Drawing.Size(41, 16);
            this.opttxt.TabIndex = 1;
            this.opttxt.Text = "Txt";
            this.opttxt.UseVisualStyleBackColor = true;
            this.opttxt.Visible = false;
            // 
            // optExcel
            // 
            this.optExcel.AutoSize = true;
            this.optExcel.Checked = true;
            this.optExcel.Location = new System.Drawing.Point(18, 19);
            this.optExcel.Name = "optExcel";
            this.optExcel.Size = new System.Drawing.Size(53, 16);
            this.optExcel.TabIndex = 0;
            this.optExcel.TabStop = true;
            this.optExcel.Text = "Excel";
            this.optExcel.UseVisualStyleBackColor = true;
            // 
            // cmdGetData
            // 
            this.cmdGetData.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.cmdGetData.Location = new System.Drawing.Point(707, 17);
            this.cmdGetData.Name = "cmdGetData";
            this.cmdGetData.Size = new System.Drawing.Size(75, 36);
            this.cmdGetData.TabIndex = 4;
            this.cmdGetData.Text = "GetData";
            this.cmdGetData.UseVisualStyleBackColor = false;
            this.cmdGetData.Click += new System.EventHandler(this.cmdGetData_Click);
            // 
            // labelInfor
            // 
            this.labelInfor.AutoSize = true;
            this.labelInfor.BackColor = System.Drawing.Color.LightPink;
            this.labelInfor.Font = new System.Drawing.Font("宋体", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.labelInfor.ForeColor = System.Drawing.SystemColors.ControlText;
            this.labelInfor.Location = new System.Drawing.Point(788, 26);
            this.labelInfor.Name = "labelInfor";
            this.labelInfor.Size = new System.Drawing.Size(97, 20);
            this.labelInfor.TabIndex = 5;
            this.labelInfor.Text = "        ";
            // 
            // FraCompN
            // 
            this.FraCompN.BackColor = System.Drawing.SystemColors.Control;
            this.FraCompN.Controls.Add(this.lblInfor);
            this.FraCompN.Controls.Add(this.DTPEndDate);
            this.FraCompN.Controls.Add(this.DTPBeginDate);
            this.FraCompN.Controls.Add(this.DTPBeginTime);
            this.FraCompN.Controls.Add(this.DTPEndTime);
            this.FraCompN.Controls.Add(this.txtVendorCode);
            this.FraCompN.Controls.Add(this.txtLotCode);
            this.FraCompN.Controls.Add(this.txtDateCode);
            this.FraCompN.Controls.Add(this.txtModel);
            this.FraCompN.Controls.Add(this.TxtCompPN);
            this.FraCompN.Controls.Add(this.label9);
            this.FraCompN.Controls.Add(this.label8);
            this.FraCompN.Controls.Add(this.label7);
            this.FraCompN.Controls.Add(this.label6);
            this.FraCompN.Controls.Add(this.label5);
            this.FraCompN.Controls.Add(this.label4);
            this.FraCompN.Controls.Add(this.label3);
            this.FraCompN.Location = new System.Drawing.Point(280, 59);
            this.FraCompN.Name = "FraCompN";
            this.FraCompN.Size = new System.Drawing.Size(605, 319);
            this.FraCompN.TabIndex = 6;
            this.FraCompN.TabStop = false;
            this.FraCompN.Text = "FraCompN";
            // 
            // lblInfor
            // 
            this.lblInfor.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.lblInfor.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblInfor.Location = new System.Drawing.Point(8, 157);
            this.lblInfor.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblInfor.Name = "lblInfor";
            this.lblInfor.Size = new System.Drawing.Size(592, 159);
            this.lblInfor.TabIndex = 21;
            this.lblInfor.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // DTPEndDate
            // 
            this.DTPEndDate.Location = new System.Drawing.Point(131, 133);
            this.DTPEndDate.Name = "DTPEndDate";
            this.DTPEndDate.Size = new System.Drawing.Size(121, 21);
            this.DTPEndDate.TabIndex = 20;
            // 
            // DTPBeginDate
            // 
            this.DTPBeginDate.Location = new System.Drawing.Point(131, 106);
            this.DTPBeginDate.Name = "DTPBeginDate";
            this.DTPBeginDate.Size = new System.Drawing.Size(121, 21);
            this.DTPBeginDate.TabIndex = 19;
            // 
            // DTPBeginTime
            // 
            this.DTPBeginTime.Location = new System.Drawing.Point(256, 104);
            this.DTPBeginTime.Name = "DTPBeginTime";
            this.DTPBeginTime.Size = new System.Drawing.Size(121, 21);
            this.DTPBeginTime.TabIndex = 17;
            // 
            // DTPEndTime
            // 
            this.DTPEndTime.Location = new System.Drawing.Point(256, 133);
            this.DTPEndTime.Name = "DTPEndTime";
            this.DTPEndTime.Size = new System.Drawing.Size(121, 21);
            this.DTPEndTime.TabIndex = 16;
            // 
            // txtVendorCode
            // 
            this.txtVendorCode.Location = new System.Drawing.Point(131, 46);
            this.txtVendorCode.Name = "txtVendorCode";
            this.txtVendorCode.Size = new System.Drawing.Size(121, 21);
            this.txtVendorCode.TabIndex = 13;
            // 
            // txtLotCode
            // 
            this.txtLotCode.Location = new System.Drawing.Point(131, 76);
            this.txtLotCode.Name = "txtLotCode";
            this.txtLotCode.Size = new System.Drawing.Size(121, 21);
            this.txtLotCode.TabIndex = 12;
            // 
            // txtDateCode
            // 
            this.txtDateCode.Location = new System.Drawing.Point(370, 46);
            this.txtDateCode.Name = "txtDateCode";
            this.txtDateCode.Size = new System.Drawing.Size(121, 21);
            this.txtDateCode.TabIndex = 11;
            // 
            // txtModel
            // 
            this.txtModel.Location = new System.Drawing.Point(370, 74);
            this.txtModel.Name = "txtModel";
            this.txtModel.Size = new System.Drawing.Size(121, 21);
            this.txtModel.TabIndex = 10;
            // 
            // TxtCompPN
            // 
            this.TxtCompPN.Location = new System.Drawing.Point(131, 16);
            this.TxtCompPN.Name = "TxtCompPN";
            this.TxtCompPN.Size = new System.Drawing.Size(121, 21);
            this.TxtCompPN.TabIndex = 9;
            // 
            // label9
            // 
            this.label9.BackColor = System.Drawing.SystemColors.Info;
            this.label9.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label9.Location = new System.Drawing.Point(6, 48);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(130, 19);
            this.label9.TabIndex = 8;
            this.label9.Text = "VendorCode:";
            // 
            // label8
            // 
            this.label8.BackColor = System.Drawing.SystemColors.Info;
            this.label8.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label8.Location = new System.Drawing.Point(6, 76);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(130, 19);
            this.label8.TabIndex = 7;
            this.label8.Text = "  Lot Code:";
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.SystemColors.Info;
            this.label7.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label7.Location = new System.Drawing.Point(6, 106);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(130, 19);
            this.label7.TabIndex = 6;
            this.label7.Text = "Begin Date:";
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.SystemColors.Info;
            this.label6.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.Location = new System.Drawing.Point(6, 134);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(130, 19);
            this.label6.TabIndex = 5;
            this.label6.Text = "  End Date:";
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.SystemColors.Info;
            this.label5.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.Location = new System.Drawing.Point(258, 48);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(119, 19);
            this.label5.TabIndex = 4;
            this.label5.Text = "Date Code:";
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.SystemColors.Info;
            this.label4.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(258, 76);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(119, 19);
            this.label4.TabIndex = 3;
            this.label4.Text = "  Model  :";
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.SystemColors.Info;
            this.label3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(6, 17);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(130, 19);
            this.label3.TabIndex = 2;
            this.label3.Text = "  CompPN  :";
            // 
            // groupBox4
            // 
            this.groupBox4.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox4.Controls.Add(this.dtData);
            this.groupBox4.Controls.Add(this.inputSN);
            this.groupBox4.Controls.Add(this.CMDChosefile);
            this.groupBox4.Controls.Add(this.Txtpath);
            this.groupBox4.Controls.Add(this.label10);
            this.groupBox4.Location = new System.Drawing.Point(280, 384);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(605, 161);
            this.groupBox4.TabIndex = 7;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "导入数据";
            // 
            // dtData
            // 
            this.dtData.BackgroundColor = System.Drawing.Color.White;
            this.dtData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dtData.Location = new System.Drawing.Point(6, 39);
            this.dtData.Name = "dtData";
            this.dtData.RowTemplate.Height = 23;
            this.dtData.Size = new System.Drawing.Size(589, 116);
            this.dtData.TabIndex = 17;
            // 
            // inputSN
            // 
            this.inputSN.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.inputSN.Location = new System.Drawing.Point(508, 11);
            this.inputSN.Name = "inputSN";
            this.inputSN.Size = new System.Drawing.Size(87, 25);
            this.inputSN.TabIndex = 16;
            this.inputSN.Text = "导入数据";
            this.inputSN.UseVisualStyleBackColor = false;
            this.inputSN.Click += new System.EventHandler(this.inputSN_Click);
            // 
            // CMDChosefile
            // 
            this.CMDChosefile.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.CMDChosefile.Location = new System.Drawing.Point(415, 11);
            this.CMDChosefile.Name = "CMDChosefile";
            this.CMDChosefile.Size = new System.Drawing.Size(87, 25);
            this.CMDChosefile.TabIndex = 15;
            this.CMDChosefile.Text = "选择文件";
            this.CMDChosefile.UseVisualStyleBackColor = false;
            this.CMDChosefile.Click += new System.EventHandler(this.CMDChosefile_Click);
            // 
            // Txtpath
            // 
            this.Txtpath.Location = new System.Drawing.Point(119, 15);
            this.Txtpath.Name = "Txtpath";
            this.Txtpath.Size = new System.Drawing.Size(290, 21);
            this.Txtpath.TabIndex = 14;
            // 
            // label10
            // 
            this.label10.BackColor = System.Drawing.SystemColors.Info;
            this.label10.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label10.Location = new System.Drawing.Point(6, 17);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(119, 19);
            this.label10.TabIndex = 8;
            this.label10.Text = "File Path:";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.label11.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label11.Location = new System.Drawing.Point(12, 548);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(472, 16);
            this.label11.TabIndex = 8;
            this.label11.Text = "如果选用Txt格式: 数据结果本地存放路径:C:\\TraceReportData\\ ";
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.label12.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label12.Location = new System.Drawing.Point(13, 573);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(440, 16);
            this.label12.TabIndex = 9;
            this.label12.Text = "数据结果服务器共享路径:QSMS Server D:\\TraceReportData\\";
            // 
            // frmTraceReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(890, 607);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.FraCompN);
            this.Controls.Add(this.labelInfor);
            this.Controls.Add(this.cmdGetData);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.CbbDataType);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.groupBox1);
            this.Name = "frmTraceReport";
            this.Text = "frmTraceReport";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmTraceReport_FormClosed);
            this.Load += new System.EventHandler(this.frmTraceReport_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DataGridSN)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.FraCompN.ResumeLayout(false);
            this.FraCompN.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dtData)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DataGridView DataGridSN;
        private System.Windows.Forms.TextBox TxtSN;
        private System.Windows.Forms.Label lblSNWO;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton OptBatch;
        private System.Windows.Forms.RadioButton OptSN;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox CbbDataType;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.RadioButton opttxt;
        private System.Windows.Forms.RadioButton optExcel;
        private System.Windows.Forms.Button cmdGetData;
        private System.Windows.Forms.Label labelInfor;
        private System.Windows.Forms.GroupBox FraCompN;
        private System.Windows.Forms.TextBox DTPBeginTime;
        private System.Windows.Forms.TextBox DTPEndTime;
        private System.Windows.Forms.TextBox txtVendorCode;
        private System.Windows.Forms.TextBox txtLotCode;
        private System.Windows.Forms.TextBox txtDateCode;
        private System.Windows.Forms.TextBox txtModel;
        private System.Windows.Forms.TextBox TxtCompPN;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button inputSN;
        private System.Windows.Forms.Button CMDChosefile;
        private System.Windows.Forms.TextBox Txtpath;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.DateTimePicker DTPEndDate;
        private System.Windows.Forms.DateTimePicker DTPBeginDate;
        private System.Windows.Forms.Label lblInfor;
        private System.Windows.Forms.DataGridView dtData;
    }
}