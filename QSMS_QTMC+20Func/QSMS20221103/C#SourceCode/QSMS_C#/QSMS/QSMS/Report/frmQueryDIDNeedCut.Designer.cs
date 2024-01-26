namespace QSMS.QSMS.Report
{
    partial class frmQueryDIDNeedCut
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtPN = new System.Windows.Forms.TextBox();
            this.cmdGetFile = new System.Windows.Forms.Button();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cboSheetName = new System.Windows.Forms.ComboBox();
            this.cmdLoad = new System.Windows.Forms.Button();
            this.lstPN = new System.Windows.Forms.ListBox();
            this.cmdClear = new System.Windows.Forms.Button();
            this.cmdQuery = new System.Windows.Forms.Button();
            this.cmdExcel = new System.Windows.Forms.Button();
            this.cmdExist = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.dtTo = new System.Windows.Forms.DateTimePicker();
            this.dtFrom = new System.Windows.Forms.DateTimePicker();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("宋体", 12F);
            this.label1.Location = new System.Drawing.Point(29, 21);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(128, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "CompomentPNList";
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.Info;
            this.label2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(29, 64);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(79, 23);
            this.label2.TabIndex = 1;
            this.label2.Text = "Comp PN";
            // 
            // txtPN
            // 
            this.txtPN.BackColor = System.Drawing.Color.PaleTurquoise;
            this.txtPN.Font = new System.Drawing.Font("宋体", 11F);
            this.txtPN.Location = new System.Drawing.Point(112, 64);
            this.txtPN.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtPN.Name = "txtPN";
            this.txtPN.Size = new System.Drawing.Size(196, 24);
            this.txtPN.TabIndex = 2;
            this.txtPN.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtPN_KeyDown);
            // 
            // cmdGetFile
            // 
            this.cmdGetFile.BackColor = System.Drawing.SystemColors.Info;
            this.cmdGetFile.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdGetFile.Location = new System.Drawing.Point(32, 104);
            this.cmdGetFile.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.cmdGetFile.Name = "cmdGetFile";
            this.cmdGetFile.Size = new System.Drawing.Size(76, 23);
            this.cmdGetFile.TabIndex = 3;
            this.cmdGetFile.Text = "...";
            this.cmdGetFile.UseVisualStyleBackColor = false;
            this.cmdGetFile.Click += new System.EventHandler(this.cmdGetFile_Click);
            // 
            // txtFilePath
            // 
            this.txtFilePath.Font = new System.Drawing.Font("宋体", 11F);
            this.txtFilePath.Location = new System.Drawing.Point(112, 105);
            this.txtFilePath.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.Size = new System.Drawing.Size(196, 24);
            this.txtFilePath.TabIndex = 4;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.SystemColors.Info;
            this.label3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(29, 147);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(79, 23);
            this.label3.TabIndex = 5;
            this.label3.Text = " Sheet ";
            // 
            // cboSheetName
            // 
            this.cboSheetName.BackColor = System.Drawing.Color.PaleTurquoise;
            this.cboSheetName.Font = new System.Drawing.Font("宋体", 11F);
            this.cboSheetName.FormattingEnabled = true;
            this.cboSheetName.Location = new System.Drawing.Point(112, 146);
            this.cboSheetName.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.cboSheetName.Name = "cboSheetName";
            this.cboSheetName.Size = new System.Drawing.Size(138, 23);
            this.cboSheetName.TabIndex = 6;
            // 
            // cmdLoad
            // 
            this.cmdLoad.Font = new System.Drawing.Font("宋体", 11F);
            this.cmdLoad.Location = new System.Drawing.Point(254, 146);
            this.cmdLoad.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.cmdLoad.Name = "cmdLoad";
            this.cmdLoad.Size = new System.Drawing.Size(53, 21);
            this.cmdLoad.TabIndex = 7;
            this.cmdLoad.Text = "Load";
            this.cmdLoad.UseVisualStyleBackColor = true;
            this.cmdLoad.Click += new System.EventHandler(this.cmdLoad_Click);
            // 
            // lstPN
            // 
            this.lstPN.BackColor = System.Drawing.SystemColors.Control;
            this.lstPN.Font = new System.Drawing.Font("宋体", 26F);
            this.lstPN.FormattingEnabled = true;
            this.lstPN.ItemHeight = 35;
            this.lstPN.Location = new System.Drawing.Point(32, 187);
            this.lstPN.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.lstPN.Name = "lstPN";
            this.lstPN.Size = new System.Drawing.Size(276, 284);
            this.lstPN.TabIndex = 8;
            // 
            // cmdClear
            // 
            this.cmdClear.BackColor = System.Drawing.Color.Tan;
            this.cmdClear.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdClear.Location = new System.Drawing.Point(334, 64);
            this.cmdClear.Margin = new System.Windows.Forms.Padding(0, 2, 0, 2);
            this.cmdClear.Name = "cmdClear";
            this.cmdClear.Size = new System.Drawing.Size(69, 30);
            this.cmdClear.TabIndex = 9;
            this.cmdClear.Text = "Clear";
            this.cmdClear.UseVisualStyleBackColor = false;
            this.cmdClear.Click += new System.EventHandler(this.cmdClear_Click);
            // 
            // cmdQuery
            // 
            this.cmdQuery.BackColor = System.Drawing.Color.Tan;
            this.cmdQuery.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdQuery.Location = new System.Drawing.Point(407, 64);
            this.cmdQuery.Margin = new System.Windows.Forms.Padding(0, 2, 0, 2);
            this.cmdQuery.Name = "cmdQuery";
            this.cmdQuery.Size = new System.Drawing.Size(68, 30);
            this.cmdQuery.TabIndex = 10;
            this.cmdQuery.Text = "Query";
            this.cmdQuery.UseVisualStyleBackColor = false;
            this.cmdQuery.Click += new System.EventHandler(this.cmdQuery_Click);
            // 
            // cmdExcel
            // 
            this.cmdExcel.BackColor = System.Drawing.Color.Tan;
            this.cmdExcel.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdExcel.Location = new System.Drawing.Point(482, 64);
            this.cmdExcel.Margin = new System.Windows.Forms.Padding(0, 2, 0, 2);
            this.cmdExcel.Name = "cmdExcel";
            this.cmdExcel.Size = new System.Drawing.Size(67, 30);
            this.cmdExcel.TabIndex = 11;
            this.cmdExcel.Text = "Excel";
            this.cmdExcel.UseVisualStyleBackColor = false;
            this.cmdExcel.Click += new System.EventHandler(this.cmdExcel_Click);
            // 
            // cmdExist
            // 
            this.cmdExist.BackColor = System.Drawing.Color.Tan;
            this.cmdExist.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdExist.Location = new System.Drawing.Point(556, 64);
            this.cmdExist.Margin = new System.Windows.Forms.Padding(0, 2, 0, 2);
            this.cmdExist.Name = "cmdExist";
            this.cmdExist.Size = new System.Drawing.Size(76, 30);
            this.cmdExist.TabIndex = 12;
            this.cmdExist.Text = "Exist";
            this.cmdExist.UseVisualStyleBackColor = false;
            this.cmdExist.Click += new System.EventHandler(this.cmdExist_Click);
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.SystemColors.Info;
            this.label4.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(331, 109);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(49, 24);
            this.label4.TabIndex = 13;
            this.label4.Text = "From:";
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.SystemColors.Info;
            this.label5.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.Location = new System.Drawing.Point(535, 109);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(38, 24);
            this.label5.TabIndex = 14;
            this.label5.Text = "To:";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.Control;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(334, 146);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowTemplate.Height = 27;
            this.dataGridView1.Size = new System.Drawing.Size(440, 322);
            this.dataGridView1.TabIndex = 17;
            // 
            // dtTo
            // 
            this.dtTo.CalendarFont = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dtTo.Font = new System.Drawing.Font("宋体", 11F);
            this.dtTo.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtTo.Location = new System.Drawing.Point(577, 109);
            this.dtTo.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.dtTo.Name = "dtTo";
            this.dtTo.Size = new System.Drawing.Size(147, 24);
            this.dtTo.TabIndex = 77;
            this.dtTo.Value = new System.DateTime(2021, 2, 22, 0, 0, 0, 0);
            // 
            // dtFrom
            // 
            this.dtFrom.CalendarFont = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dtFrom.Font = new System.Drawing.Font("宋体", 11F);
            this.dtFrom.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtFrom.Location = new System.Drawing.Point(384, 109);
            this.dtFrom.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.dtFrom.Name = "dtFrom";
            this.dtFrom.Size = new System.Drawing.Size(140, 24);
            this.dtFrom.TabIndex = 76;
            this.dtFrom.Value = new System.DateTime(2021, 2, 22, 0, 0, 0, 0);
            // 
            // frmQueryDIDNeedCut
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(790, 509);
            this.Controls.Add(this.dtTo);
            this.Controls.Add(this.dtFrom);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cmdExist);
            this.Controls.Add(this.cmdExcel);
            this.Controls.Add(this.cmdQuery);
            this.Controls.Add(this.cmdClear);
            this.Controls.Add(this.lstPN);
            this.Controls.Add(this.cmdLoad);
            this.Controls.Add(this.cboSheetName);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtFilePath);
            this.Controls.Add(this.cmdGetFile);
            this.Controls.Add(this.txtPN);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "frmQueryDIDNeedCut";
            this.Text = "Query DID Need Cut";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmQueryDIDNeedCut_FormClosed);
            this.Load += new System.EventHandler(this.frmQueryDIDNeedCut_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtPN;
        private System.Windows.Forms.Button cmdGetFile;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cboSheetName;
        private System.Windows.Forms.Button cmdLoad;
        private System.Windows.Forms.ListBox lstPN;
        private System.Windows.Forms.Button cmdClear;
        private System.Windows.Forms.Button cmdQuery;
        private System.Windows.Forms.Button cmdExcel;
        private System.Windows.Forms.Button cmdExist;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DateTimePicker dtTo;
        private System.Windows.Forms.DateTimePicker dtFrom;
    }
}