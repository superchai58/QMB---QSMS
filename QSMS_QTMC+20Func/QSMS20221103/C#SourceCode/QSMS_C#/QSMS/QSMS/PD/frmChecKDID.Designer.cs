namespace QSMS.QSMS.PD
{
    partial class frmChecKDID
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
            this.DTPBeginDate = new System.Windows.Forms.DateTimePicker();
            this.lblBeginT = new System.Windows.Forms.Label();
            this.DTPEndDate = new System.Windows.Forms.DateTimePicker();
            this.lblEndT = new System.Windows.Forms.Label();
            this.txtBeginT = new System.Windows.Forms.TextBox();
            this.txtEndT = new System.Windows.Forms.TextBox();
            this.txtDID = new System.Windows.Forms.TextBox();
            this.lblDID = new System.Windows.Forms.Label();
            this.txtBarCode = new System.Windows.Forms.TextBox();
            this.lblBarCode = new System.Windows.Forms.Label();
            this.cboSide = new System.Windows.Forms.ComboBox();
            this.lblSide = new System.Windows.Forms.Label();
            this.lblLine = new System.Windows.Forms.Label();
            this.cboLine = new System.Windows.Forms.ComboBox();
            this.lblGroupID = new System.Windows.Forms.Label();
            this.cboGroupID = new System.Windows.Forms.ComboBox();
            this.lblTabel = new System.Windows.Forms.Label();
            this.cboTabel = new System.Windows.Forms.ComboBox();
            this.btnQuery = new System.Windows.Forms.Button();
            this.btnExcel = new System.Windows.Forms.Button();
            this.DG_Result = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.DG_Result)).BeginInit();
            this.SuspendLayout();
            // 
            // DTPBeginDate
            // 
            this.DTPBeginDate.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.DTPBeginDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.DTPBeginDate.Location = new System.Drawing.Point(99, 12);
            this.DTPBeginDate.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.DTPBeginDate.Name = "DTPBeginDate";
            this.DTPBeginDate.Size = new System.Drawing.Size(125, 26);
            this.DTPBeginDate.TabIndex = 2;
            this.DTPBeginDate.Value = new System.DateTime(2019, 11, 22, 12, 41, 44, 0);
            // 
            // lblBeginT
            // 
            this.lblBeginT.BackColor = System.Drawing.SystemColors.Info;
            this.lblBeginT.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblBeginT.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblBeginT.Location = new System.Drawing.Point(9, 13);
            this.lblBeginT.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblBeginT.Name = "lblBeginT";
            this.lblBeginT.Size = new System.Drawing.Size(86, 24);
            this.lblBeginT.TabIndex = 1;
            this.lblBeginT.Text = "BeginDate:";
            this.lblBeginT.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // DTPEndDate
            // 
            this.DTPEndDate.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.DTPEndDate.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.DTPEndDate.Location = new System.Drawing.Point(99, 50);
            this.DTPEndDate.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.DTPEndDate.Name = "DTPEndDate";
            this.DTPEndDate.Size = new System.Drawing.Size(125, 26);
            this.DTPEndDate.TabIndex = 5;
            this.DTPEndDate.Value = new System.DateTime(2019, 11, 22, 12, 41, 44, 0);
            // 
            // lblEndT
            // 
            this.lblEndT.BackColor = System.Drawing.SystemColors.Info;
            this.lblEndT.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblEndT.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblEndT.Location = new System.Drawing.Point(9, 50);
            this.lblEndT.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblEndT.Name = "lblEndT";
            this.lblEndT.Size = new System.Drawing.Size(86, 24);
            this.lblEndT.TabIndex = 4;
            this.lblEndT.Text = "EndDate:";
            this.lblEndT.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtBeginT
            // 
            this.txtBeginT.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtBeginT.Location = new System.Drawing.Point(227, 12);
            this.txtBeginT.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtBeginT.Name = "txtBeginT";
            this.txtBeginT.Size = new System.Drawing.Size(98, 26);
            this.txtBeginT.TabIndex = 3;
            // 
            // txtEndT
            // 
            this.txtEndT.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtEndT.Location = new System.Drawing.Point(227, 49);
            this.txtEndT.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtEndT.Name = "txtEndT";
            this.txtEndT.Size = new System.Drawing.Size(98, 26);
            this.txtEndT.TabIndex = 6;
            // 
            // txtDID
            // 
            this.txtDID.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtDID.Location = new System.Drawing.Point(99, 89);
            this.txtDID.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtDID.Name = "txtDID";
            this.txtDID.Size = new System.Drawing.Size(226, 26);
            this.txtDID.TabIndex = 8;
            this.txtDID.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDID_KeyDown);
            // 
            // lblDID
            // 
            this.lblDID.BackColor = System.Drawing.SystemColors.Info;
            this.lblDID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblDID.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblDID.Location = new System.Drawing.Point(9, 89);
            this.lblDID.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblDID.Name = "lblDID";
            this.lblDID.Size = new System.Drawing.Size(86, 24);
            this.lblDID.TabIndex = 7;
            this.lblDID.Text = "DID:";
            this.lblDID.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtBarCode
            // 
            this.txtBarCode.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtBarCode.Location = new System.Drawing.Point(99, 126);
            this.txtBarCode.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtBarCode.Name = "txtBarCode";
            this.txtBarCode.Size = new System.Drawing.Size(226, 26);
            this.txtBarCode.TabIndex = 10;
            this.txtBarCode.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtBarCode_KeyDown);
            // 
            // lblBarCode
            // 
            this.lblBarCode.BackColor = System.Drawing.SystemColors.Info;
            this.lblBarCode.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblBarCode.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblBarCode.Location = new System.Drawing.Point(9, 126);
            this.lblBarCode.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblBarCode.Name = "lblBarCode";
            this.lblBarCode.Size = new System.Drawing.Size(86, 24);
            this.lblBarCode.TabIndex = 9;
            this.lblBarCode.Text = "BarCode:";
            this.lblBarCode.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // cboSide
            // 
            this.cboSide.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cboSide.FormattingEnabled = true;
            this.cboSide.Location = new System.Drawing.Point(454, 13);
            this.cboSide.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.cboSide.Name = "cboSide";
            this.cboSide.Size = new System.Drawing.Size(158, 24);
            this.cboSide.TabIndex = 12;
            // 
            // lblSide
            // 
            this.lblSide.BackColor = System.Drawing.SystemColors.Info;
            this.lblSide.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblSide.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblSide.Location = new System.Drawing.Point(363, 13);
            this.lblSide.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblSide.Name = "lblSide";
            this.lblSide.Size = new System.Drawing.Size(86, 24);
            this.lblSide.TabIndex = 11;
            this.lblSide.Text = "Side:";
            this.lblSide.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblLine
            // 
            this.lblLine.BackColor = System.Drawing.SystemColors.Info;
            this.lblLine.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblLine.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblLine.Location = new System.Drawing.Point(363, 50);
            this.lblLine.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblLine.Name = "lblLine";
            this.lblLine.Size = new System.Drawing.Size(86, 24);
            this.lblLine.TabIndex = 13;
            this.lblLine.Text = "Line:";
            this.lblLine.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // cboLine
            // 
            this.cboLine.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cboLine.FormattingEnabled = true;
            this.cboLine.Location = new System.Drawing.Point(454, 50);
            this.cboLine.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.cboLine.Name = "cboLine";
            this.cboLine.Size = new System.Drawing.Size(158, 24);
            this.cboLine.TabIndex = 14;
            this.cboLine.DropDownClosed += new System.EventHandler(this.cboLine_DropDownClosed);
            // 
            // lblGroupID
            // 
            this.lblGroupID.BackColor = System.Drawing.SystemColors.Info;
            this.lblGroupID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblGroupID.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblGroupID.Location = new System.Drawing.Point(363, 89);
            this.lblGroupID.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblGroupID.Name = "lblGroupID";
            this.lblGroupID.Size = new System.Drawing.Size(86, 24);
            this.lblGroupID.TabIndex = 15;
            this.lblGroupID.Text = "GroupID:";
            this.lblGroupID.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // cboGroupID
            // 
            this.cboGroupID.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cboGroupID.FormattingEnabled = true;
            this.cboGroupID.Location = new System.Drawing.Point(454, 89);
            this.cboGroupID.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.cboGroupID.Name = "cboGroupID";
            this.cboGroupID.Size = new System.Drawing.Size(158, 24);
            this.cboGroupID.TabIndex = 16;
            // 
            // lblTabel
            // 
            this.lblTabel.BackColor = System.Drawing.SystemColors.Info;
            this.lblTabel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTabel.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblTabel.Location = new System.Drawing.Point(363, 126);
            this.lblTabel.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblTabel.Name = "lblTabel";
            this.lblTabel.Size = new System.Drawing.Size(86, 24);
            this.lblTabel.TabIndex = 17;
            this.lblTabel.Text = "Tabel:";
            this.lblTabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // cboTabel
            // 
            this.cboTabel.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cboTabel.FormattingEnabled = true;
            this.cboTabel.Location = new System.Drawing.Point(454, 126);
            this.cboTabel.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.cboTabel.Name = "cboTabel";
            this.cboTabel.Size = new System.Drawing.Size(158, 24);
            this.cboTabel.TabIndex = 18;
            // 
            // btnQuery
            // 
            this.btnQuery.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnQuery.Location = new System.Drawing.Point(637, 46);
            this.btnQuery.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnQuery.Name = "btnQuery";
            this.btnQuery.Size = new System.Drawing.Size(75, 28);
            this.btnQuery.TabIndex = 19;
            this.btnQuery.Text = "Query";
            this.btnQuery.UseVisualStyleBackColor = true;
            this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // btnExcel
            // 
            this.btnExcel.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnExcel.Location = new System.Drawing.Point(637, 108);
            this.btnExcel.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(75, 28);
            this.btnExcel.TabIndex = 20;
            this.btnExcel.Text = "Excel";
            this.btnExcel.UseVisualStyleBackColor = true;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // DG_Result
            // 
            this.DG_Result.AllowUserToAddRows = false;
            this.DG_Result.BackgroundColor = System.Drawing.SystemColors.ButtonFace;
            this.DG_Result.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DG_Result.Location = new System.Drawing.Point(10, 165);
            this.DG_Result.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.DG_Result.Name = "DG_Result";
            this.DG_Result.ReadOnly = true;
            this.DG_Result.RowTemplate.Height = 27;
            this.DG_Result.Size = new System.Drawing.Size(702, 295);
            this.DG_Result.TabIndex = 21;
            // 
            // frmChecKDID
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(737, 481);
            this.Controls.Add(this.DG_Result);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.btnQuery);
            this.Controls.Add(this.lblTabel);
            this.Controls.Add(this.cboTabel);
            this.Controls.Add(this.lblGroupID);
            this.Controls.Add(this.cboGroupID);
            this.Controls.Add(this.lblLine);
            this.Controls.Add(this.cboLine);
            this.Controls.Add(this.lblSide);
            this.Controls.Add(this.cboSide);
            this.Controls.Add(this.txtBarCode);
            this.Controls.Add(this.lblBarCode);
            this.Controls.Add(this.txtDID);
            this.Controls.Add(this.lblDID);
            this.Controls.Add(this.txtEndT);
            this.Controls.Add(this.txtBeginT);
            this.Controls.Add(this.DTPEndDate);
            this.Controls.Add(this.lblEndT);
            this.Controls.Add(this.DTPBeginDate);
            this.Controls.Add(this.lblBeginT);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "frmChecKDID";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ChecKDID[20201130]";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmChecKDID_FormClosed);
            this.Load += new System.EventHandler(this.frmChecKDID_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DG_Result)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DateTimePicker DTPBeginDate;
        private System.Windows.Forms.Label lblBeginT;
        private System.Windows.Forms.DateTimePicker DTPEndDate;
        private System.Windows.Forms.Label lblEndT;
        private System.Windows.Forms.TextBox txtBeginT;
        private System.Windows.Forms.TextBox txtEndT;
        private System.Windows.Forms.TextBox txtDID;
        private System.Windows.Forms.Label lblDID;
        private System.Windows.Forms.TextBox txtBarCode;
        private System.Windows.Forms.Label lblBarCode;
        private System.Windows.Forms.ComboBox cboSide;
        private System.Windows.Forms.Label lblSide;
        private System.Windows.Forms.Label lblLine;
        private System.Windows.Forms.ComboBox cboLine;
        private System.Windows.Forms.Label lblGroupID;
        private System.Windows.Forms.ComboBox cboGroupID;
        private System.Windows.Forms.Label lblTabel;
        private System.Windows.Forms.ComboBox cboTabel;
        private System.Windows.Forms.Button btnQuery;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.DataGridView DG_Result;
    }
}