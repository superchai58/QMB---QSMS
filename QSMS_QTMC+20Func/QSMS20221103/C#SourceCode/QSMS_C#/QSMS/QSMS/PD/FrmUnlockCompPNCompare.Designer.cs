namespace QSMS.QSMS.PD
{
    partial class FrmUnlockCompPNCompare
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
            this.txtReason = new System.Windows.Forms.TextBox();
            this.txtDID = new System.Windows.Forms.TextBox();
            this.txtCompPN = new System.Windows.Forms.TextBox();
            this.lblStatus = new System.Windows.Forms.TextBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.GV_Data1 = new System.Windows.Forms.DataGridView();
            this.btnUnlock = new System.Windows.Forms.Button();
            this.textBox12 = new System.Windows.Forms.TextBox();
            this.textBox11 = new System.Windows.Forms.TextBox();
            this.textBox10 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.CobLine = new System.Windows.Forms.ComboBox();
            this.GV_Data2 = new System.Windows.Forms.DataGridView();
            this.btnQuery = new System.Windows.Forms.Button();
            this.btnExcel = new System.Windows.Forms.Button();
            this.dtpSDate = new System.Windows.Forms.DateTimePicker();
            this.dtpEDate = new System.Windows.Forms.DateTimePicker();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.GV_Data1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.GV_Data2)).BeginInit();
            this.SuspendLayout();
            // 
            // txtReason
            // 
            this.txtReason.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtReason.Location = new System.Drawing.Point(122, 91);
            this.txtReason.Multiline = true;
            this.txtReason.Name = "txtReason";
            this.txtReason.Size = new System.Drawing.Size(290, 79);
            this.txtReason.TabIndex = 2;
            // 
            // txtDID
            // 
            this.txtDID.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtDID.Location = new System.Drawing.Point(122, 22);
            this.txtDID.Name = "txtDID";
            this.txtDID.Size = new System.Drawing.Size(290, 26);
            this.txtDID.TabIndex = 3;
            // 
            // txtCompPN
            // 
            this.txtCompPN.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtCompPN.Location = new System.Drawing.Point(122, 57);
            this.txtCompPN.Name = "txtCompPN";
            this.txtCompPN.Size = new System.Drawing.Size(290, 26);
            this.txtCompPN.TabIndex = 5;
            this.txtCompPN.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCompPN_KeyPress);
            // 
            // lblStatus
            // 
            this.lblStatus.BackColor = System.Drawing.Color.Bisque;
            this.lblStatus.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.lblStatus.Location = new System.Drawing.Point(122, 194);
            this.lblStatus.Multiline = true;
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.ReadOnly = true;
            this.lblStatus.Size = new System.Drawing.Size(290, 41);
            this.lblStatus.TabIndex = 8;
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.GV_Data1);
            this.panel1.Controls.Add(this.btnUnlock);
            this.panel1.Controls.Add(this.textBox12);
            this.panel1.Controls.Add(this.textBox11);
            this.panel1.Controls.Add(this.textBox10);
            this.panel1.Controls.Add(this.lblStatus);
            this.panel1.Controls.Add(this.txtReason);
            this.panel1.Controls.Add(this.txtCompPN);
            this.panel1.Controls.Add(this.txtDID);
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(901, 250);
            this.panel1.TabIndex = 9;
            this.panel1.Tag = "Unlock";
            // 
            // GV_Data1
            // 
            this.GV_Data1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.GV_Data1.Location = new System.Drawing.Point(503, 23);
            this.GV_Data1.Name = "GV_Data1";
            this.GV_Data1.RowTemplate.Height = 23;
            this.GV_Data1.Size = new System.Drawing.Size(384, 212);
            this.GV_Data1.TabIndex = 14;
            this.GV_Data1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.GV_Data1_CellClick);
            // 
            // btnUnlock
            // 
            this.btnUnlock.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnUnlock.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnUnlock.Location = new System.Drawing.Point(429, 23);
            this.btnUnlock.Name = "btnUnlock";
            this.btnUnlock.Size = new System.Drawing.Size(68, 73);
            this.btnUnlock.TabIndex = 13;
            this.btnUnlock.Text = "Unlock";
            this.btnUnlock.UseVisualStyleBackColor = false;
            this.btnUnlock.Click += new System.EventHandler(this.btnUnlock_Click);
            // 
            // textBox12
            // 
            this.textBox12.BackColor = System.Drawing.SystemColors.Info;
            this.textBox12.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox12.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox12.Location = new System.Drawing.Point(16, 58);
            this.textBox12.Name = "textBox12";
            this.textBox12.ReadOnly = true;
            this.textBox12.Size = new System.Drawing.Size(100, 19);
            this.textBox12.TabIndex = 12;
            this.textBox12.Text = "CompPN";
            // 
            // textBox11
            // 
            this.textBox11.BackColor = System.Drawing.SystemColors.Info;
            this.textBox11.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox11.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox11.Location = new System.Drawing.Point(16, 92);
            this.textBox11.Name = "textBox11";
            this.textBox11.ReadOnly = true;
            this.textBox11.Size = new System.Drawing.Size(100, 19);
            this.textBox11.TabIndex = 11;
            this.textBox11.Text = "Reason";
            // 
            // textBox10
            // 
            this.textBox10.BackColor = System.Drawing.SystemColors.Info;
            this.textBox10.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox10.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox10.Location = new System.Drawing.Point(16, 23);
            this.textBox10.Name = "textBox10";
            this.textBox10.ReadOnly = true;
            this.textBox10.Size = new System.Drawing.Size(100, 19);
            this.textBox10.TabIndex = 10;
            this.textBox10.Text = "DID";
            // 
            // textBox2
            // 
            this.textBox2.BackColor = System.Drawing.SystemColors.Info;
            this.textBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox2.Location = new System.Drawing.Point(245, 265);
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.Size = new System.Drawing.Size(100, 19);
            this.textBox2.TabIndex = 12;
            this.textBox2.Text = "BeginDate";
            // 
            // textBox4
            // 
            this.textBox4.BackColor = System.Drawing.SystemColors.Info;
            this.textBox4.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox4.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox4.Location = new System.Drawing.Point(17, 265);
            this.textBox4.Name = "textBox4";
            this.textBox4.ReadOnly = true;
            this.textBox4.Size = new System.Drawing.Size(100, 19);
            this.textBox4.TabIndex = 13;
            this.textBox4.Text = "Line";
            // 
            // textBox5
            // 
            this.textBox5.BackColor = System.Drawing.SystemColors.Info;
            this.textBox5.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox5.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox5.Location = new System.Drawing.Point(536, 264);
            this.textBox5.Name = "textBox5";
            this.textBox5.ReadOnly = true;
            this.textBox5.Size = new System.Drawing.Size(100, 19);
            this.textBox5.TabIndex = 14;
            this.textBox5.Text = "End Date";
            // 
            // CobLine
            // 
            this.CobLine.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CobLine.FormattingEnabled = true;
            this.CobLine.Location = new System.Drawing.Point(123, 263);
            this.CobLine.Name = "CobLine";
            this.CobLine.Size = new System.Drawing.Size(88, 24);
            this.CobLine.TabIndex = 15;
            // 
            // GV_Data2
            // 
            this.GV_Data2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.GV_Data2.Location = new System.Drawing.Point(17, 303);
            this.GV_Data2.Name = "GV_Data2";
            this.GV_Data2.RowTemplate.Height = 23;
            this.GV_Data2.Size = new System.Drawing.Size(772, 215);
            this.GV_Data2.TabIndex = 15;
            // 
            // btnQuery
            // 
            this.btnQuery.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnQuery.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnQuery.Location = new System.Drawing.Point(815, 303);
            this.btnQuery.Name = "btnQuery";
            this.btnQuery.Size = new System.Drawing.Size(75, 63);
            this.btnQuery.TabIndex = 16;
            this.btnQuery.Text = "Query";
            this.btnQuery.UseVisualStyleBackColor = false;
            this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // btnExcel
            // 
            this.btnExcel.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnExcel.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnExcel.Location = new System.Drawing.Point(813, 387);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(75, 60);
            this.btnExcel.TabIndex = 17;
            this.btnExcel.Text = "Excel";
            this.btnExcel.UseVisualStyleBackColor = false;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // dtpSDate
            // 
            this.dtpSDate.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dtpSDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpSDate.Location = new System.Drawing.Point(351, 261);
            this.dtpSDate.Name = "dtpSDate";
            this.dtpSDate.Size = new System.Drawing.Size(147, 26);
            this.dtpSDate.TabIndex = 18;
            // 
            // dtpEDate
            // 
            this.dtpEDate.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dtpEDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpEDate.Location = new System.Drawing.Point(642, 261);
            this.dtpEDate.Name = "dtpEDate";
            this.dtpEDate.Size = new System.Drawing.Size(147, 26);
            this.dtpEDate.TabIndex = 19;
            // 
            // FrmUnlockCompPNCompare
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(902, 530);
            this.Controls.Add(this.dtpEDate);
            this.Controls.Add(this.dtpSDate);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.btnQuery);
            this.Controls.Add(this.GV_Data2);
            this.Controls.Add(this.CobLine);
            this.Controls.Add(this.textBox5);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.panel1);
            this.Name = "FrmUnlockCompPNCompare";
            this.Text = "UnlockCompPNCompare";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FrmUnlockCompPNCompare_FormClosed);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.GV_Data1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.GV_Data2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtReason;
        private System.Windows.Forms.TextBox txtDID;
        private System.Windows.Forms.TextBox txtCompPN;
        private System.Windows.Forms.TextBox lblStatus;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox textBox12;
        private System.Windows.Forms.TextBox textBox11;
        private System.Windows.Forms.TextBox textBox10;
        private System.Windows.Forms.DataGridView GV_Data1;
        private System.Windows.Forms.Button btnUnlock;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.TextBox textBox5;
        private System.Windows.Forms.ComboBox CobLine;
        private System.Windows.Forms.DataGridView GV_Data2;
        private System.Windows.Forms.Button btnQuery;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.DateTimePicker dtpSDate;
        private System.Windows.Forms.DateTimePicker dtpEDate;
    }
}