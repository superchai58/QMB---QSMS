namespace QSMS.QSMS.MCC
{
    partial class frmDIDCheckStock
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
            this.cmdQuery = new System.Windows.Forms.Button();
            this.chkStock = new System.Windows.Forms.CheckBox();
            this.lblRefID = new System.Windows.Forms.Label();
            this.txtRefID = new System.Windows.Forms.TextBox();
            this.cmdExcel = new System.Windows.Forms.Button();
            this.lblMsg = new System.Windows.Forms.Label();
            this.dataGrid1 = new System.Windows.Forms.DataGridView();
            this.dataGrid2 = new System.Windows.Forms.DataGridView();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.dtpEDate = new System.Windows.Forms.DateTimePicker();
            this.dtpSDate = new System.Windows.Forms.DateTimePicker();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid2)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(25, 32);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(95, 26);
            this.label1.TabIndex = 0;
            this.label1.Text = "BeginDate";
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.Info;
            this.label2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(310, 32);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(98, 23);
            this.label2.TabIndex = 2;
            this.label2.Text = "EndDate";
            // 
            // cmdQuery
            // 
            this.cmdQuery.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.cmdQuery.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdQuery.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.cmdQuery.Location = new System.Drawing.Point(702, 23);
            this.cmdQuery.Margin = new System.Windows.Forms.Padding(2);
            this.cmdQuery.Name = "cmdQuery";
            this.cmdQuery.Size = new System.Drawing.Size(86, 35);
            this.cmdQuery.TabIndex = 5;
            this.cmdQuery.Text = "Find";
            this.cmdQuery.UseVisualStyleBackColor = false;
            this.cmdQuery.Click += new System.EventHandler(this.cmdQuery_Click);
            // 
            // chkStock
            // 
            this.chkStock.BackColor = System.Drawing.SystemColors.Info;
            this.chkStock.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.chkStock.Location = new System.Drawing.Point(28, 72);
            this.chkStock.Margin = new System.Windows.Forms.Padding(2);
            this.chkStock.Name = "chkStock";
            this.chkStock.Size = new System.Drawing.Size(248, 23);
            this.chkStock.TabIndex = 6;
            this.chkStock.Text = "Check Qty with MC Stock";
            this.chkStock.UseVisualStyleBackColor = false;
            // 
            // lblRefID
            // 
            this.lblRefID.BackColor = System.Drawing.SystemColors.Info;
            this.lblRefID.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblRefID.Location = new System.Drawing.Point(310, 72);
            this.lblRefID.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblRefID.Name = "lblRefID";
            this.lblRefID.Size = new System.Drawing.Size(98, 23);
            this.lblRefID.TabIndex = 7;
            this.lblRefID.Text = "ReferenceID";
            // 
            // txtRefID
            // 
            this.txtRefID.Location = new System.Drawing.Point(412, 74);
            this.txtRefID.Margin = new System.Windows.Forms.Padding(2);
            this.txtRefID.Name = "txtRefID";
            this.txtRefID.Size = new System.Drawing.Size(260, 21);
            this.txtRefID.TabIndex = 8;
            this.txtRefID.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRefID_KeyPress);
            // 
            // cmdExcel
            // 
            this.cmdExcel.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.cmdExcel.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdExcel.Location = new System.Drawing.Point(703, 65);
            this.cmdExcel.Margin = new System.Windows.Forms.Padding(2);
            this.cmdExcel.Name = "cmdExcel";
            this.cmdExcel.Size = new System.Drawing.Size(85, 35);
            this.cmdExcel.TabIndex = 9;
            this.cmdExcel.Text = "Excel";
            this.cmdExcel.UseVisualStyleBackColor = false;
            this.cmdExcel.Click += new System.EventHandler(this.cmdExcel_Click);
            // 
            // lblMsg
            // 
            this.lblMsg.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.lblMsg.Font = new System.Drawing.Font("宋体", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblMsg.Location = new System.Drawing.Point(25, 110);
            this.lblMsg.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(1011, 42);
            this.lblMsg.TabIndex = 10;
            this.lblMsg.Text = "Message:";
            this.lblMsg.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // dataGrid1
            // 
            this.dataGrid1.AllowUserToAddRows = false;
            this.dataGrid1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGrid1.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this.dataGrid1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGrid1.Location = new System.Drawing.Point(28, 182);
            this.dataGrid1.Margin = new System.Windows.Forms.Padding(2);
            this.dataGrid1.Name = "dataGrid1";
            this.dataGrid1.RowTemplate.Height = 44;
            this.dataGrid1.Size = new System.Drawing.Size(493, 417);
            this.dataGrid1.TabIndex = 11;
            // 
            // dataGrid2
            // 
            this.dataGrid2.AllowUserToAddRows = false;
            this.dataGrid2.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dataGrid2.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this.dataGrid2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGrid2.Location = new System.Drawing.Point(546, 182);
            this.dataGrid2.Margin = new System.Windows.Forms.Padding(2);
            this.dataGrid2.Name = "dataGrid2";
            this.dataGrid2.RowTemplate.Height = 44;
            this.dataGrid2.Size = new System.Drawing.Size(490, 417);
            this.dataGrid2.TabIndex = 12;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.SystemColors.Info;
            this.label3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(25, 161);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(80, 16);
            this.label3.TabIndex = 13;
            this.label3.Text = "CompPN OK";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.SystemColors.Info;
            this.label4.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(543, 161);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(112, 16);
            this.label4.TabIndex = 14;
            this.label4.Text = "CompPN Not OK";
            // 
            // dtpEDate
            // 
            this.dtpEDate.CalendarFont = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dtpEDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpEDate.Location = new System.Drawing.Point(412, 32);
            this.dtpEDate.Margin = new System.Windows.Forms.Padding(2);
            this.dtpEDate.Name = "dtpEDate";
            this.dtpEDate.Size = new System.Drawing.Size(260, 21);
            this.dtpEDate.TabIndex = 75;
            this.dtpEDate.Value = new System.DateTime(2021, 9, 23, 0, 0, 0, 0);
            // 
            // dtpSDate
            // 
            this.dtpSDate.CalendarFont = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dtpSDate.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dtpSDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpSDate.Location = new System.Drawing.Point(124, 32);
            this.dtpSDate.Margin = new System.Windows.Forms.Padding(2);
            this.dtpSDate.Name = "dtpSDate";
            this.dtpSDate.Size = new System.Drawing.Size(152, 26);
            this.dtpSDate.TabIndex = 74;
            this.dtpSDate.Value = new System.DateTime(2021, 2, 22, 0, 0, 0, 0);
            // 
            // frmDIDCheckStock
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1059, 610);
            this.Controls.Add(this.dtpEDate);
            this.Controls.Add(this.dtpSDate);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.dataGrid2);
            this.Controls.Add(this.dataGrid1);
            this.Controls.Add(this.lblMsg);
            this.Controls.Add(this.cmdExcel);
            this.Controls.Add(this.txtRefID);
            this.Controls.Add(this.lblRefID);
            this.Controls.Add(this.chkStock);
            this.Controls.Add(this.cmdQuery);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "frmDIDCheckStock";
            this.Text = "Check DID Stock 20210923";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmDIDCheckStock_FormClosed);
            this.Load += new System.EventHandler(this.frmDIDCheckStock_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGrid2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button cmdQuery;
        private System.Windows.Forms.CheckBox chkStock;
        private System.Windows.Forms.Label lblRefID;
        private System.Windows.Forms.TextBox txtRefID;
        private System.Windows.Forms.Button cmdExcel;
        private System.Windows.Forms.Label lblMsg;
        private System.Windows.Forms.DataGridView dataGrid1;
        private System.Windows.Forms.DataGridView dataGrid2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DateTimePicker dtpEDate;
        private System.Windows.Forms.DateTimePicker dtpSDate;
    }
}