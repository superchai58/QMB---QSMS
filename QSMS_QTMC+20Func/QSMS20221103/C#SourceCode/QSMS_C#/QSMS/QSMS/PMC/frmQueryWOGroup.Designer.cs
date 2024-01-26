namespace QSMS.QSMS.PMC
{
    partial class frmQueryWOGroup
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
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.cboLine = new System.Windows.Forms.ComboBox();
            this.txtWO = new System.Windows.Forms.TextBox();
            this.cmdQuery = new System.Windows.Forms.Button();
            this.DGNotFinished = new System.Windows.Forms.DataGridView();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.DGFinish = new System.Windows.Forms.DataGridView();
            this.cmdClosed = new System.Windows.Forms.Button();
            this.cmdUnClosed = new System.Windows.Forms.Button();
            this.dtpEDate = new System.Windows.Forms.DateTimePicker();
            this.dtpSDate = new System.Windows.Forms.DateTimePicker();
            ((System.ComponentModel.ISupportInitialize)(this.DGNotFinished)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGFinish)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(16, 9);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 23);
            this.label1.TabIndex = 0;
            this.label1.Text = "Line:";
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.Info;
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(16, 40);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(100, 23);
            this.label2.TabIndex = 1;
            this.label2.Text = "BeginDate:";
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.SystemColors.Info;
            this.label3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(16, 71);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(100, 23);
            this.label3.TabIndex = 2;
            this.label3.Text = "EndDate:";
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.SystemColors.Info;
            this.label4.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.label4.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(16, 102);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(100, 23);
            this.label4.TabIndex = 3;
            this.label4.Text = "WorkOrder:";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cboLine
            // 
            this.cboLine.Font = new System.Drawing.Font("宋体", 11F);
            this.cboLine.FormattingEnabled = true;
            this.cboLine.Location = new System.Drawing.Point(130, 9);
            this.cboLine.Margin = new System.Windows.Forms.Padding(2);
            this.cboLine.Name = "cboLine";
            this.cboLine.Size = new System.Drawing.Size(173, 23);
            this.cboLine.TabIndex = 4;
            // 
            // txtWO
            // 
            this.txtWO.Font = new System.Drawing.Font("宋体", 11F);
            this.txtWO.Location = new System.Drawing.Point(130, 101);
            this.txtWO.Margin = new System.Windows.Forms.Padding(2);
            this.txtWO.Name = "txtWO";
            this.txtWO.Size = new System.Drawing.Size(173, 24);
            this.txtWO.TabIndex = 7;
            // 
            // cmdQuery
            // 
            this.cmdQuery.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.cmdQuery.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdQuery.Location = new System.Drawing.Point(326, 53);
            this.cmdQuery.Margin = new System.Windows.Forms.Padding(2);
            this.cmdQuery.Name = "cmdQuery";
            this.cmdQuery.Size = new System.Drawing.Size(110, 39);
            this.cmdQuery.TabIndex = 8;
            this.cmdQuery.Text = "QueryWO";
            this.cmdQuery.UseVisualStyleBackColor = false;
            this.cmdQuery.Click += new System.EventHandler(this.cmdQuery_Click);
            // 
            // DGNotFinished
            // 
            this.DGNotFinished.AllowUserToAddRows = false;
            this.DGNotFinished.AllowUserToDeleteRows = false;
            this.DGNotFinished.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.DGNotFinished.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            this.DGNotFinished.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGNotFinished.Location = new System.Drawing.Point(16, 158);
            this.DGNotFinished.Margin = new System.Windows.Forms.Padding(2, 2, 2, 12);
            this.DGNotFinished.Name = "DGNotFinished";
            this.DGNotFinished.ReadOnly = true;
            this.DGNotFinished.RowTemplate.Height = 44;
            this.DGNotFinished.Size = new System.Drawing.Size(808, 183);
            this.DGNotFinished.TabIndex = 9;
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Font = new System.Drawing.Font("微软雅黑 Light", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox1.Location = new System.Drawing.Point(15, 138);
            this.textBox1.Margin = new System.Windows.Forms.Padding(2);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(644, 16);
            this.textBox1.TabIndex = 10;
            this.textBox1.Text = "UnClosedWOList";
            this.textBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox2
            // 
            this.textBox2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.textBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox2.Font = new System.Drawing.Font("微软雅黑 Light", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox2.Location = new System.Drawing.Point(16, 355);
            this.textBox2.Margin = new System.Windows.Forms.Padding(2);
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.Size = new System.Drawing.Size(643, 16);
            this.textBox2.TabIndex = 11;
            this.textBox2.Text = "ClosedWOList";
            this.textBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // DGFinish
            // 
            this.DGFinish.AllowUserToAddRows = false;
            this.DGFinish.AllowUserToDeleteRows = false;
            this.DGFinish.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.DGFinish.BackgroundColor = System.Drawing.SystemColors.ButtonHighlight;
            this.DGFinish.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGFinish.Location = new System.Drawing.Point(15, 375);
            this.DGFinish.Margin = new System.Windows.Forms.Padding(2);
            this.DGFinish.Name = "DGFinish";
            this.DGFinish.RowTemplate.Height = 44;
            this.DGFinish.Size = new System.Drawing.Size(808, 207);
            this.DGFinish.TabIndex = 12;
            // 
            // cmdClosed
            // 
            this.cmdClosed.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.cmdClosed.Font = new System.Drawing.Font("微软雅黑 Light", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdClosed.Location = new System.Drawing.Point(693, 344);
            this.cmdClosed.Margin = new System.Windows.Forms.Padding(2);
            this.cmdClosed.Name = "cmdClosed";
            this.cmdClosed.Size = new System.Drawing.Size(105, 27);
            this.cmdClosed.TabIndex = 13;
            this.cmdClosed.Text = "Excel Closed";
            this.cmdClosed.UseVisualStyleBackColor = false;
            this.cmdClosed.Click += new System.EventHandler(this.cmdClosed_Click);
            // 
            // cmdUnClosed
            // 
            this.cmdUnClosed.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.cmdUnClosed.Font = new System.Drawing.Font("微软雅黑 Light", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdUnClosed.Location = new System.Drawing.Point(693, 125);
            this.cmdUnClosed.Margin = new System.Windows.Forms.Padding(2);
            this.cmdUnClosed.Name = "cmdUnClosed";
            this.cmdUnClosed.Size = new System.Drawing.Size(105, 29);
            this.cmdUnClosed.TabIndex = 14;
            this.cmdUnClosed.Text = "Excel UnClosed";
            this.cmdUnClosed.UseVisualStyleBackColor = false;
            this.cmdUnClosed.Click += new System.EventHandler(this.cmdUnClosed_Click);
            // 
            // dtpEDate
            // 
            this.dtpEDate.CalendarFont = new System.Drawing.Font("宋体", 11F);
            this.dtpEDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpEDate.Location = new System.Drawing.Point(130, 71);
            this.dtpEDate.Margin = new System.Windows.Forms.Padding(2);
            this.dtpEDate.Name = "dtpEDate";
            this.dtpEDate.Size = new System.Drawing.Size(173, 21);
            this.dtpEDate.TabIndex = 77;
            this.dtpEDate.Value = new System.DateTime(2021, 2, 22, 0, 0, 0, 0);
            // 
            // dtpSDate
            // 
            this.dtpSDate.CalendarFont = new System.Drawing.Font("宋体", 11F);
            this.dtpSDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpSDate.Location = new System.Drawing.Point(130, 40);
            this.dtpSDate.Margin = new System.Windows.Forms.Padding(2);
            this.dtpSDate.Name = "dtpSDate";
            this.dtpSDate.Size = new System.Drawing.Size(173, 21);
            this.dtpSDate.TabIndex = 76;
            this.dtpSDate.Value = new System.DateTime(2021, 2, 22, 0, 0, 0, 0);
            // 
            // frmQueryWOGroup
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(834, 593);
            this.Controls.Add(this.dtpEDate);
            this.Controls.Add(this.dtpSDate);
            this.Controls.Add(this.cmdUnClosed);
            this.Controls.Add(this.cmdClosed);
            this.Controls.Add(this.DGFinish);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.DGNotFinished);
            this.Controls.Add(this.cmdQuery);
            this.Controls.Add(this.txtWO);
            this.Controls.Add(this.cboLine);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "frmQueryWOGroup";
            this.Text = "QueryWOGroup";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmQueryWOGroup_FormClosed);
            this.Load += new System.EventHandler(this.frmQueryWOGroup_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DGNotFinished)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.DGFinish)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cboLine;
        private System.Windows.Forms.TextBox txtWO;
        private System.Windows.Forms.Button cmdQuery;
        private System.Windows.Forms.DataGridView DGNotFinished;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.DataGridView DGFinish;
        private System.Windows.Forms.Button cmdClosed;
        private System.Windows.Forms.Button cmdUnClosed;
        private System.Windows.Forms.DateTimePicker dtpEDate;
        private System.Windows.Forms.DateTimePicker dtpSDate;
    }
}