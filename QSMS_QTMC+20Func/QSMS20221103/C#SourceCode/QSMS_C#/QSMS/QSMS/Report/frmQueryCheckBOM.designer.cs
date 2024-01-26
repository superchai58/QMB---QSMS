namespace QSMS.QSMS.Report
{
    partial class frmQueryCheckBOM
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
            this.CmdQuery = new System.Windows.Forms.Button();
            this.Command1 = new System.Windows.Forms.Button();
            this.TxtWO = new System.Windows.Forms.TextBox();
            this.CboLine = new System.Windows.Forms.ComboBox();
            this.dtpEDate = new System.Windows.Forms.DateTimePicker();
            this.dtpSDate = new System.Windows.Forms.DateTimePicker();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.DG1 = new System.Windows.Forms.DataGridView();
            this.label5 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DG1)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.TxtWO);
            this.groupBox1.Controls.Add(this.CboLine);
            this.groupBox1.Controls.Add(this.dtpEDate);
            this.groupBox1.Controls.Add(this.dtpSDate);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.groupBox1.Location = new System.Drawing.Point(2, 0);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox1.Size = new System.Drawing.Size(555, 108);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "MainTain";
            // 
            // CmdQuery
            // 
            this.CmdQuery.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.CmdQuery.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CmdQuery.Location = new System.Drawing.Point(574, 14);
            this.CmdQuery.Margin = new System.Windows.Forms.Padding(2);
            this.CmdQuery.Name = "CmdQuery";
            this.CmdQuery.Size = new System.Drawing.Size(74, 34);
            this.CmdQuery.TabIndex = 20;
            this.CmdQuery.Text = "Query";
            this.CmdQuery.UseVisualStyleBackColor = false;
            this.CmdQuery.Click += new System.EventHandler(this.CmdQuery_Click);
            // 
            // Command1
            // 
            this.Command1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.Command1.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Command1.Location = new System.Drawing.Point(574, 58);
            this.Command1.Margin = new System.Windows.Forms.Padding(2);
            this.Command1.Name = "Command1";
            this.Command1.Size = new System.Drawing.Size(74, 34);
            this.Command1.TabIndex = 19;
            this.Command1.Text = "Exit";
            this.Command1.UseVisualStyleBackColor = false;
            this.Command1.Click += new System.EventHandler(this.Command1_Click);
            // 
            // TxtWO
            // 
            this.TxtWO.Location = new System.Drawing.Point(98, 69);
            this.TxtWO.Margin = new System.Windows.Forms.Padding(2);
            this.TxtWO.Name = "TxtWO";
            this.TxtWO.Size = new System.Drawing.Size(163, 23);
            this.TxtWO.TabIndex = 18;
            // 
            // CboLine
            // 
            this.CboLine.FormattingEnabled = true;
            this.CboLine.Location = new System.Drawing.Point(368, 69);
            this.CboLine.Margin = new System.Windows.Forms.Padding(2);
            this.CboLine.Name = "CboLine";
            this.CboLine.Size = new System.Drawing.Size(163, 25);
            this.CboLine.TabIndex = 17;
            // 
            // dtpEDate
            // 
            this.dtpEDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpEDate.Location = new System.Drawing.Point(368, 26);
            this.dtpEDate.Margin = new System.Windows.Forms.Padding(2);
            this.dtpEDate.Name = "dtpEDate";
            this.dtpEDate.Size = new System.Drawing.Size(163, 23);
            this.dtpEDate.TabIndex = 16;
            this.dtpEDate.Value = new System.DateTime(2020, 11, 25, 0, 0, 0, 0);
            // 
            // dtpSDate
            // 
            this.dtpSDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpSDate.Location = new System.Drawing.Point(98, 25);
            this.dtpSDate.Margin = new System.Windows.Forms.Padding(2);
            this.dtpSDate.Name = "dtpSDate";
            this.dtpSDate.Size = new System.Drawing.Size(163, 23);
            this.dtpSDate.TabIndex = 15;
            this.dtpSDate.Value = new System.DateTime(2020, 11, 25, 0, 0, 0, 0);
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.SystemColors.Info;
            this.label4.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(289, 26);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(75, 22);
            this.label4.TabIndex = 3;
            this.label4.Text = "End Date";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.SystemColors.Info;
            this.label3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(286, 69);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(78, 22);
            this.label3.TabIndex = 2;
            this.label3.Text = "Line";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.Info;
            this.label2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(4, 68);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(89, 22);
            this.label2.TabIndex = 1;
            this.label2.Text = "Work_date";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(7, 26);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(86, 22);
            this.label1.TabIndex = 0;
            this.label1.Text = "BeginDate";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // DG1
            // 
            this.DG1.BackgroundColor = System.Drawing.SystemColors.Control;
            this.DG1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DG1.Location = new System.Drawing.Point(2, 131);
            this.DG1.Margin = new System.Windows.Forms.Padding(2);
            this.DG1.Name = "DG1";
            this.DG1.RowTemplate.Height = 27;
            this.DG1.Size = new System.Drawing.Size(646, 260);
            this.DG1.TabIndex = 1;
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.SystemColors.Info;
            this.label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label5.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.Location = new System.Drawing.Point(2, 110);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(646, 19);
            this.label5.TabIndex = 2;
            this.label5.Text = "All work order are check bom fail";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // frmQueryCheckBOM
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(671, 402);
            this.Controls.Add(this.Command1);
            this.Controls.Add(this.CmdQuery);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.DG1);
            this.Controls.Add(this.groupBox1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "frmQueryCheckBOM";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Query wo check bom information[2009-11-18]";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmQueryCheckBOM_FormClosed);
            this.Load += new System.EventHandler(this.FrmQueryCheckBOM_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.DG1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button CmdQuery;
        private System.Windows.Forms.Button Command1;
        private System.Windows.Forms.TextBox TxtWO;
        private System.Windows.Forms.ComboBox CboLine;
        private System.Windows.Forms.DateTimePicker dtpEDate;
        private System.Windows.Forms.DateTimePicker dtpSDate;
        private System.Windows.Forms.DataGridView DG1;
        private System.Windows.Forms.Label label5;
    }
}