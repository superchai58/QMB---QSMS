namespace QSMS.QSMS.IPQC
{
    partial class frmInspection_Del
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
            this.txtDID = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dtpEDateTime = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.dtpSDateTime = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.btnQuery = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.txtTransdatetime = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.txtLower = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtResult = new System.Windows.Forms.TextBox();
            this.txtValue = new System.Windows.Forms.TextBox();
            this.txtUpper = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.txtCompPN = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.btnDel = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // txtDID
            // 
            this.txtDID.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtDID.Location = new System.Drawing.Point(111, 19);
            this.txtDID.Margin = new System.Windows.Forms.Padding(2);
            this.txtDID.Name = "txtDID";
            this.txtDID.Size = new System.Drawing.Size(425, 26);
            this.txtDID.TabIndex = 25;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(8, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(98, 26);
            this.label1.TabIndex = 26;
            this.label1.Text = "DID\\CompPN:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dtpEDateTime);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.dtpSDateTime);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.btnQuery);
            this.groupBox1.Controls.Add(this.txtDID);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(9, 10);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox1.Size = new System.Drawing.Size(680, 97);
            this.groupBox1.TabIndex = 27;
            this.groupBox1.TabStop = false;
            // 
            // dtpEDateTime
            // 
            this.dtpEDateTime.CalendarFont = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dtpEDateTime.CustomFormat = "yyyy-MM-dd HH:mm:ss";
            this.dtpEDateTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpEDateTime.Location = new System.Drawing.Point(377, 64);
            this.dtpEDateTime.Margin = new System.Windows.Forms.Padding(2);
            this.dtpEDateTime.Name = "dtpEDateTime";
            this.dtpEDateTime.ShowUpDown = true;
            this.dtpEDateTime.Size = new System.Drawing.Size(158, 21);
            this.dtpEDateTime.TabIndex = 32;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.SystemColors.Info;
            this.label3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(274, 59);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(98, 26);
            this.label3.TabIndex = 31;
            this.label3.Text = "EndDate:";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // dtpSDateTime
            // 
            this.dtpSDateTime.CalendarFont = new System.Drawing.Font("微软雅黑", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dtpSDateTime.CustomFormat = "yyyy-MM-dd HH:mm:ss";
            this.dtpSDateTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpSDateTime.Location = new System.Drawing.Point(111, 65);
            this.dtpSDateTime.Margin = new System.Windows.Forms.Padding(2);
            this.dtpSDateTime.Name = "dtpSDateTime";
            this.dtpSDateTime.ShowUpDown = true;
            this.dtpSDateTime.Size = new System.Drawing.Size(158, 21);
            this.dtpSDateTime.TabIndex = 30;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.Info;
            this.label2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(8, 60);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(98, 26);
            this.label2.TabIndex = 29;
            this.label2.Text = "StartDate:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnQuery
            // 
            this.btnQuery.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnQuery.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnQuery.Location = new System.Drawing.Point(565, 19);
            this.btnQuery.Margin = new System.Windows.Forms.Padding(2);
            this.btnQuery.Name = "btnQuery";
            this.btnQuery.Size = new System.Drawing.Size(70, 42);
            this.btnQuery.TabIndex = 28;
            this.btnQuery.Text = "Query";
            this.btnQuery.UseVisualStyleBackColor = false;
            this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.txtTransdatetime);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.txtLower);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.txtResult);
            this.groupBox2.Controls.Add(this.txtValue);
            this.groupBox2.Controls.Add(this.txtUpper);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.txtCompPN);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Location = new System.Drawing.Point(9, 113);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox2.Size = new System.Drawing.Size(680, 97);
            this.groupBox2.TabIndex = 33;
            this.groupBox2.TabStop = false;
            // 
            // txtTransdatetime
            // 
            this.txtTransdatetime.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.txtTransdatetime.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtTransdatetime.Location = new System.Drawing.Point(551, 54);
            this.txtTransdatetime.Margin = new System.Windows.Forms.Padding(2);
            this.txtTransdatetime.Name = "txtTransdatetime";
            this.txtTransdatetime.ReadOnly = true;
            this.txtTransdatetime.Size = new System.Drawing.Size(99, 26);
            this.txtTransdatetime.TabIndex = 39;
            // 
            // label9
            // 
            this.label9.BackColor = System.Drawing.SystemColors.Info;
            this.label9.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label9.Location = new System.Drawing.Point(424, 54);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(122, 26);
            this.label9.TabIndex = 38;
            this.label9.Text = "TransDateTime:";
            this.label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label8
            // 
            this.label8.BackColor = System.Drawing.SystemColors.Info;
            this.label8.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label8.Location = new System.Drawing.Point(468, 18);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(63, 26);
            this.label8.TabIndex = 37;
            this.label8.Text = "Lower:";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtLower
            // 
            this.txtLower.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.txtLower.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtLower.Location = new System.Drawing.Point(536, 18);
            this.txtLower.Margin = new System.Windows.Forms.Padding(2);
            this.txtLower.Name = "txtLower";
            this.txtLower.ReadOnly = true;
            this.txtLower.Size = new System.Drawing.Size(113, 26);
            this.txtLower.TabIndex = 36;
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.SystemColors.Info;
            this.label7.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label7.Location = new System.Drawing.Point(283, 18);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(63, 26);
            this.label7.TabIndex = 35;
            this.label7.Text = "Upper:";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtResult
            // 
            this.txtResult.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.txtResult.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtResult.Location = new System.Drawing.Point(338, 55);
            this.txtResult.Margin = new System.Windows.Forms.Padding(2);
            this.txtResult.Name = "txtResult";
            this.txtResult.ReadOnly = true;
            this.txtResult.Size = new System.Drawing.Size(82, 26);
            this.txtResult.TabIndex = 34;
            // 
            // txtValue
            // 
            this.txtValue.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.txtValue.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtValue.Location = new System.Drawing.Point(111, 55);
            this.txtValue.Margin = new System.Windows.Forms.Padding(2);
            this.txtValue.Name = "txtValue";
            this.txtValue.ReadOnly = true;
            this.txtValue.Size = new System.Drawing.Size(118, 26);
            this.txtValue.TabIndex = 33;
            // 
            // txtUpper
            // 
            this.txtUpper.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.txtUpper.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtUpper.Location = new System.Drawing.Point(353, 19);
            this.txtUpper.Margin = new System.Windows.Forms.Padding(2);
            this.txtUpper.Name = "txtUpper";
            this.txtUpper.ReadOnly = true;
            this.txtUpper.Size = new System.Drawing.Size(110, 26);
            this.txtUpper.TabIndex = 32;
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.SystemColors.Info;
            this.label4.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(233, 55);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(100, 26);
            this.label4.TabIndex = 31;
            this.label4.Text = "TestResult:";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.SystemColors.Info;
            this.label5.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.Location = new System.Drawing.Point(8, 54);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(98, 26);
            this.label5.TabIndex = 29;
            this.label5.Text = "TestValue:";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtCompPN
            // 
            this.txtCompPN.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.txtCompPN.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtCompPN.Location = new System.Drawing.Point(111, 19);
            this.txtCompPN.Margin = new System.Windows.Forms.Padding(2);
            this.txtCompPN.Name = "txtCompPN";
            this.txtCompPN.ReadOnly = true;
            this.txtCompPN.Size = new System.Drawing.Size(168, 26);
            this.txtCompPN.TabIndex = 25;
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.SystemColors.Info;
            this.label6.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.Location = new System.Drawing.Point(8, 18);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(98, 26);
            this.label6.TabIndex = 26;
            this.label6.Text = "DID\\CompPN:";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnDel
            // 
            this.btnDel.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnDel.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnDel.Location = new System.Drawing.Point(693, 122);
            this.btnDel.Margin = new System.Windows.Forms.Padding(2);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(70, 35);
            this.btnDel.TabIndex = 34;
            this.btnDel.Text = "Delete";
            this.btnDel.UseVisualStyleBackColor = false;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnCancel.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnCancel.Location = new System.Drawing.Point(693, 163);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(2);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(70, 35);
            this.btnCancel.TabIndex = 35;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = false;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            this.dataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(9, 224);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(2);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.RowTemplate.Height = 27;
            this.dataGridView1.Size = new System.Drawing.Size(754, 332);
            this.dataGridView1.TabIndex = 36;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
            this.dataGridView1.KeyUp += new System.Windows.Forms.KeyEventHandler(this.dataGridView1_KeyUp);
            // 
            // frmInspection_Del
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(809, 574);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "frmInspection_Del";
            this.Text = "Delete Inspection Result [20100421]";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmInspection_Del_FormClosed);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TextBox txtDID;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnQuery;
        private System.Windows.Forms.DateTimePicker dtpEDateTime;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker dtpSDateTime;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtCompPN;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtTransdatetime;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox txtLower;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtResult;
        private System.Windows.Forms.TextBox txtValue;
        private System.Windows.Forms.TextBox txtUpper;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.DataGridView dataGridView1;
    }
}