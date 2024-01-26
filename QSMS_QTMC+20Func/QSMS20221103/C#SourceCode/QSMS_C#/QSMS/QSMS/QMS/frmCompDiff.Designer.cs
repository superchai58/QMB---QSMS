namespace QSMS.QSMS.QMS
{
    partial class frmCompDiff
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
            this.gbSelectWorkOrder = new System.Windows.Forms.GroupBox();
            this.gbSmallBoardWO = new System.Windows.Forms.GroupBox();
            this.cboSBWO = new System.Windows.Forms.ComboBox();
            this.cboOKWO = new System.Windows.Forms.ComboBox();
            this.cboGroupID = new System.Windows.Forms.ComboBox();
            this.dtpEDate = new System.Windows.Forms.DateTimePicker();
            this.dtpBDate = new System.Windows.Forms.DateTimePicker();
            this.cboLine = new System.Windows.Forms.ComboBox();
            this.txtQty = new System.Windows.Forms.TextBox();
            this.txtRev = new System.Windows.Forms.TextBox();
            this.txtMBPN = new System.Windows.Forms.TextBox();
            this.txtModel = new System.Windows.Forms.TextBox();
            this.txtWO = new System.Windows.Forms.TextBox();
            this.btnCheck = new System.Windows.Forms.Button();
            this.btnFind = new System.Windows.Forms.Button();
            this.lblOKWorkOrder = new System.Windows.Forms.Label();
            this.lblGroupID = new System.Windows.Forms.Label();
            this.lblQty = new System.Windows.Forms.Label();
            this.lblRev = new System.Windows.Forms.Label();
            this.lblMBPN = new System.Windows.Forms.Label();
            this.lblModel = new System.Windows.Forms.Label();
            this.lblWO = new System.Windows.Forms.Label();
            this.lblLine = new System.Windows.Forms.Label();
            this.lblEndDate = new System.Windows.Forms.Label();
            this.lblBeginDate = new System.Windows.Forms.Label();
            this.rbtnGroup = new System.Windows.Forms.RadioButton();
            this.rbtnRelease = new System.Windows.Forms.RadioButton();
            this.gbDispatchQty = new System.Windows.Forms.GroupBox();
            this.dgvDispatch = new System.Windows.Forms.DataGridView();
            this.gbSelectWorkOrder.SuspendLayout();
            this.gbSmallBoardWO.SuspendLayout();
            this.gbDispatchQty.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvDispatch)).BeginInit();
            this.SuspendLayout();
            // 
            // gbSelectWorkOrder
            // 
            this.gbSelectWorkOrder.Controls.Add(this.gbSmallBoardWO);
            this.gbSelectWorkOrder.Controls.Add(this.cboOKWO);
            this.gbSelectWorkOrder.Controls.Add(this.cboGroupID);
            this.gbSelectWorkOrder.Controls.Add(this.dtpEDate);
            this.gbSelectWorkOrder.Controls.Add(this.dtpBDate);
            this.gbSelectWorkOrder.Controls.Add(this.cboLine);
            this.gbSelectWorkOrder.Controls.Add(this.txtQty);
            this.gbSelectWorkOrder.Controls.Add(this.txtRev);
            this.gbSelectWorkOrder.Controls.Add(this.txtMBPN);
            this.gbSelectWorkOrder.Controls.Add(this.txtModel);
            this.gbSelectWorkOrder.Controls.Add(this.txtWO);
            this.gbSelectWorkOrder.Controls.Add(this.btnCheck);
            this.gbSelectWorkOrder.Controls.Add(this.btnFind);
            this.gbSelectWorkOrder.Controls.Add(this.lblOKWorkOrder);
            this.gbSelectWorkOrder.Controls.Add(this.lblGroupID);
            this.gbSelectWorkOrder.Controls.Add(this.lblQty);
            this.gbSelectWorkOrder.Controls.Add(this.lblRev);
            this.gbSelectWorkOrder.Controls.Add(this.lblMBPN);
            this.gbSelectWorkOrder.Controls.Add(this.lblModel);
            this.gbSelectWorkOrder.Controls.Add(this.lblWO);
            this.gbSelectWorkOrder.Controls.Add(this.lblLine);
            this.gbSelectWorkOrder.Controls.Add(this.lblEndDate);
            this.gbSelectWorkOrder.Controls.Add(this.lblBeginDate);
            this.gbSelectWorkOrder.Controls.Add(this.rbtnGroup);
            this.gbSelectWorkOrder.Controls.Add(this.rbtnRelease);
            this.gbSelectWorkOrder.Font = new System.Drawing.Font("宋体", 9F);
            this.gbSelectWorkOrder.Location = new System.Drawing.Point(14, 12);
            this.gbSelectWorkOrder.Name = "gbSelectWorkOrder";
            this.gbSelectWorkOrder.Size = new System.Drawing.Size(831, 223);
            this.gbSelectWorkOrder.TabIndex = 0;
            this.gbSelectWorkOrder.TabStop = false;
            this.gbSelectWorkOrder.Text = "Select Work Order";
            // 
            // gbSmallBoardWO
            // 
            this.gbSmallBoardWO.Controls.Add(this.cboSBWO);
            this.gbSmallBoardWO.Location = new System.Drawing.Point(460, 84);
            this.gbSmallBoardWO.Name = "gbSmallBoardWO";
            this.gbSmallBoardWO.Size = new System.Drawing.Size(190, 45);
            this.gbSmallBoardWO.TabIndex = 24;
            this.gbSmallBoardWO.TabStop = false;
            this.gbSmallBoardWO.Text = "Small Board WO";
            this.gbSmallBoardWO.Visible = false;
            // 
            // cboSBWO
            // 
            this.cboSBWO.FormattingEnabled = true;
            this.cboSBWO.Location = new System.Drawing.Point(6, 19);
            this.cboSBWO.Name = "cboSBWO";
            this.cboSBWO.Size = new System.Drawing.Size(178, 20);
            this.cboSBWO.TabIndex = 0;
            // 
            // cboOKWO
            // 
            this.cboOKWO.FormattingEnabled = true;
            this.cboOKWO.Location = new System.Drawing.Point(472, 56);
            this.cboOKWO.Name = "cboOKWO";
            this.cboOKWO.Size = new System.Drawing.Size(178, 20);
            this.cboOKWO.TabIndex = 23;
            this.cboOKWO.SelectedValueChanged += new System.EventHandler(this.cboOKWO_SelectedValueChanged);
            // 
            // cboGroupID
            // 
            this.cboGroupID.FormattingEnabled = true;
            this.cboGroupID.Location = new System.Drawing.Point(472, 21);
            this.cboGroupID.Name = "cboGroupID";
            this.cboGroupID.Size = new System.Drawing.Size(178, 20);
            this.cboGroupID.TabIndex = 22;
            this.cboGroupID.SelectedValueChanged += new System.EventHandler(this.cboGroupID_SelectedValueChanged);
            // 
            // dtpEDate
            // 
            this.dtpEDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpEDate.Location = new System.Drawing.Point(117, 93);
            this.dtpEDate.Name = "dtpEDate";
            this.dtpEDate.Size = new System.Drawing.Size(120, 21);
            this.dtpEDate.TabIndex = 21;
            this.dtpEDate.Value = new System.DateTime(2021, 9, 2, 0, 0, 0, 0);
            // 
            // dtpBDate
            // 
            this.dtpBDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpBDate.Location = new System.Drawing.Point(116, 56);
            this.dtpBDate.Name = "dtpBDate";
            this.dtpBDate.Size = new System.Drawing.Size(121, 21);
            this.dtpBDate.TabIndex = 20;
            this.dtpBDate.Value = new System.DateTime(2021, 9, 2, 0, 0, 0, 0);
            // 
            // cboLine
            // 
            this.cboLine.Font = new System.Drawing.Font("宋体", 11.5F);
            this.cboLine.FormattingEnabled = true;
            this.cboLine.Location = new System.Drawing.Point(116, 130);
            this.cboLine.Name = "cboLine";
            this.cboLine.Size = new System.Drawing.Size(121, 23);
            this.cboLine.TabIndex = 19;
            // 
            // txtQty
            // 
            this.txtQty.BackColor = System.Drawing.Color.White;
            this.txtQty.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtQty.Location = new System.Drawing.Point(738, 175);
            this.txtQty.Name = "txtQty";
            this.txtQty.ReadOnly = true;
            this.txtQty.Size = new System.Drawing.Size(66, 21);
            this.txtQty.TabIndex = 18;
            // 
            // txtRev
            // 
            this.txtRev.BackColor = System.Drawing.Color.White;
            this.txtRev.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtRev.Location = new System.Drawing.Point(595, 175);
            this.txtRev.Name = "txtRev";
            this.txtRev.ReadOnly = true;
            this.txtRev.Size = new System.Drawing.Size(85, 21);
            this.txtRev.TabIndex = 17;
            // 
            // txtMBPN
            // 
            this.txtMBPN.BackColor = System.Drawing.Color.White;
            this.txtMBPN.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtMBPN.Location = new System.Drawing.Point(429, 175);
            this.txtMBPN.Name = "txtMBPN";
            this.txtMBPN.ReadOnly = true;
            this.txtMBPN.Size = new System.Drawing.Size(110, 21);
            this.txtMBPN.TabIndex = 16;
            // 
            // txtModel
            // 
            this.txtModel.BackColor = System.Drawing.Color.White;
            this.txtModel.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtModel.Location = new System.Drawing.Point(232, 175);
            this.txtModel.Name = "txtModel";
            this.txtModel.ReadOnly = true;
            this.txtModel.Size = new System.Drawing.Size(120, 21);
            this.txtModel.TabIndex = 15;
            // 
            // txtWO
            // 
            this.txtWO.BackColor = System.Drawing.Color.White;
            this.txtWO.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtWO.ForeColor = System.Drawing.Color.Blue;
            this.txtWO.Location = new System.Drawing.Point(80, 175);
            this.txtWO.Name = "txtWO";
            this.txtWO.ReadOnly = true;
            this.txtWO.Size = new System.Drawing.Size(89, 21);
            this.txtWO.TabIndex = 14;
            // 
            // btnCheck
            // 
            this.btnCheck.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnCheck.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnCheck.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnCheck.Location = new System.Drawing.Point(667, 97);
            this.btnCheck.Name = "btnCheck";
            this.btnCheck.Size = new System.Drawing.Size(75, 30);
            this.btnCheck.TabIndex = 13;
            this.btnCheck.Text = "Check";
            this.btnCheck.UseVisualStyleBackColor = false;
            this.btnCheck.Click += new System.EventHandler(this.btnCheck_Click);
            // 
            // btnFind
            // 
            this.btnFind.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnFind.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnFind.Location = new System.Drawing.Point(243, 65);
            this.btnFind.Name = "btnFind";
            this.btnFind.Size = new System.Drawing.Size(87, 46);
            this.btnFind.TabIndex = 12;
            this.btnFind.Text = "&Find";
            this.btnFind.UseVisualStyleBackColor = false;
            this.btnFind.Click += new System.EventHandler(this.btnFind_Click);
            // 
            // lblOKWorkOrder
            // 
            this.lblOKWorkOrder.BackColor = System.Drawing.SystemColors.Info;
            this.lblOKWorkOrder.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblOKWorkOrder.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblOKWorkOrder.Location = new System.Drawing.Point(336, 57);
            this.lblOKWorkOrder.Name = "lblOKWorkOrder";
            this.lblOKWorkOrder.Size = new System.Drawing.Size(130, 23);
            this.lblOKWorkOrder.TabIndex = 11;
            this.lblOKWorkOrder.Text = "OK Work Order:";
            this.lblOKWorkOrder.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblGroupID
            // 
            this.lblGroupID.BackColor = System.Drawing.SystemColors.Info;
            this.lblGroupID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblGroupID.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblGroupID.Location = new System.Drawing.Point(336, 20);
            this.lblGroupID.Name = "lblGroupID";
            this.lblGroupID.Size = new System.Drawing.Size(130, 23);
            this.lblGroupID.TabIndex = 10;
            this.lblGroupID.Text = "GroupID:";
            this.lblGroupID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblQty
            // 
            this.lblQty.BackColor = System.Drawing.SystemColors.Info;
            this.lblQty.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblQty.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblQty.Location = new System.Drawing.Point(686, 175);
            this.lblQty.Name = "lblQty";
            this.lblQty.Size = new System.Drawing.Size(49, 23);
            this.lblQty.TabIndex = 9;
            this.lblQty.Text = "Qty:";
            this.lblQty.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblRev
            // 
            this.lblRev.BackColor = System.Drawing.SystemColors.Info;
            this.lblRev.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblRev.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblRev.Location = new System.Drawing.Point(543, 175);
            this.lblRev.Name = "lblRev";
            this.lblRev.Size = new System.Drawing.Size(49, 23);
            this.lblRev.TabIndex = 8;
            this.lblRev.Text = "Rev:";
            this.lblRev.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblMBPN
            // 
            this.lblMBPN.BackColor = System.Drawing.SystemColors.Info;
            this.lblMBPN.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblMBPN.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblMBPN.Location = new System.Drawing.Point(358, 175);
            this.lblMBPN.Name = "lblMBPN";
            this.lblMBPN.Size = new System.Drawing.Size(70, 23);
            this.lblMBPN.TabIndex = 7;
            this.lblMBPN.Text = "MB PN:";
            this.lblMBPN.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblModel
            // 
            this.lblModel.BackColor = System.Drawing.SystemColors.Info;
            this.lblModel.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblModel.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblModel.Location = new System.Drawing.Point(170, 175);
            this.lblModel.Name = "lblModel";
            this.lblModel.Size = new System.Drawing.Size(60, 23);
            this.lblModel.TabIndex = 6;
            this.lblModel.Text = "Model:";
            this.lblModel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblWO
            // 
            this.lblWO.BackColor = System.Drawing.SystemColors.Info;
            this.lblWO.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblWO.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblWO.Location = new System.Drawing.Point(24, 175);
            this.lblWO.Name = "lblWO";
            this.lblWO.Size = new System.Drawing.Size(55, 23);
            this.lblWO.TabIndex = 5;
            this.lblWO.Text = "WO:";
            this.lblWO.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblLine
            // 
            this.lblLine.BackColor = System.Drawing.SystemColors.Info;
            this.lblLine.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblLine.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblLine.Location = new System.Drawing.Point(24, 130);
            this.lblLine.Name = "lblLine";
            this.lblLine.Size = new System.Drawing.Size(97, 23);
            this.lblLine.TabIndex = 4;
            this.lblLine.Text = "Line:";
            this.lblLine.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblEndDate
            // 
            this.lblEndDate.BackColor = System.Drawing.SystemColors.Info;
            this.lblEndDate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblEndDate.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblEndDate.Location = new System.Drawing.Point(24, 93);
            this.lblEndDate.Name = "lblEndDate";
            this.lblEndDate.Size = new System.Drawing.Size(97, 23);
            this.lblEndDate.TabIndex = 3;
            this.lblEndDate.Text = "EndDate:";
            this.lblEndDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblBeginDate
            // 
            this.lblBeginDate.BackColor = System.Drawing.SystemColors.Info;
            this.lblBeginDate.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblBeginDate.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblBeginDate.Location = new System.Drawing.Point(24, 56);
            this.lblBeginDate.Name = "lblBeginDate";
            this.lblBeginDate.Size = new System.Drawing.Size(97, 23);
            this.lblBeginDate.TabIndex = 2;
            this.lblBeginDate.Text = "BeginDate:";
            this.lblBeginDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // rbtnGroup
            // 
            this.rbtnGroup.BackColor = System.Drawing.SystemColors.Info;
            this.rbtnGroup.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rbtnGroup.Location = new System.Drawing.Point(140, 20);
            this.rbtnGroup.Name = "rbtnGroup";
            this.rbtnGroup.Size = new System.Drawing.Size(97, 24);
            this.rbtnGroup.TabIndex = 1;
            this.rbtnGroup.TabStop = true;
            this.rbtnGroup.Text = "Group";
            this.rbtnGroup.UseVisualStyleBackColor = false;
            // 
            // rbtnRelease
            // 
            this.rbtnRelease.BackColor = System.Drawing.SystemColors.Info;
            this.rbtnRelease.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rbtnRelease.Location = new System.Drawing.Point(24, 21);
            this.rbtnRelease.Name = "rbtnRelease";
            this.rbtnRelease.Size = new System.Drawing.Size(97, 24);
            this.rbtnRelease.TabIndex = 0;
            this.rbtnRelease.TabStop = true;
            this.rbtnRelease.Text = "Release";
            this.rbtnRelease.UseVisualStyleBackColor = false;
            // 
            // gbDispatchQty
            // 
            this.gbDispatchQty.Controls.Add(this.dgvDispatch);
            this.gbDispatchQty.Font = new System.Drawing.Font("宋体", 9F);
            this.gbDispatchQty.Location = new System.Drawing.Point(14, 251);
            this.gbDispatchQty.Name = "gbDispatchQty";
            this.gbDispatchQty.Size = new System.Drawing.Size(831, 346);
            this.gbDispatchQty.TabIndex = 1;
            this.gbDispatchQty.TabStop = false;
            this.gbDispatchQty.Text = "Dispatch Qty - QSMS_WO No Match Records";
            // 
            // dgvDispatch
            // 
            this.dgvDispatch.AllowUserToAddRows = false;
            this.dgvDispatch.AllowUserToDeleteRows = false;
            this.dgvDispatch.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvDispatch.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgvDispatch.Location = new System.Drawing.Point(3, 17);
            this.dgvDispatch.Name = "dgvDispatch";
            this.dgvDispatch.ReadOnly = true;
            this.dgvDispatch.RowHeadersVisible = false;
            this.dgvDispatch.RowTemplate.Height = 23;
            this.dgvDispatch.Size = new System.Drawing.Size(825, 326);
            this.dgvDispatch.TabIndex = 0;
            // 
            // frmCompDiff
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(867, 609);
            this.Controls.Add(this.gbDispatchQty);
            this.Controls.Add(this.gbSelectWorkOrder);
            this.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Name = "frmCompDiff";
            this.Text = "FrmTransferDispatchedDID";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmCompDiff_FormClosed);
            this.Load += new System.EventHandler(this.frmCompDiff_Load);
            this.gbSelectWorkOrder.ResumeLayout(false);
            this.gbSelectWorkOrder.PerformLayout();
            this.gbSmallBoardWO.ResumeLayout(false);
            this.gbDispatchQty.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvDispatch)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox gbSelectWorkOrder;
        private System.Windows.Forms.GroupBox gbSmallBoardWO;
        private System.Windows.Forms.ComboBox cboSBWO;
        private System.Windows.Forms.ComboBox cboOKWO;
        private System.Windows.Forms.ComboBox cboGroupID;
        private System.Windows.Forms.DateTimePicker dtpEDate;
        private System.Windows.Forms.DateTimePicker dtpBDate;
        private System.Windows.Forms.ComboBox cboLine;
        private System.Windows.Forms.TextBox txtQty;
        private System.Windows.Forms.TextBox txtRev;
        private System.Windows.Forms.TextBox txtMBPN;
        private System.Windows.Forms.TextBox txtModel;
        private System.Windows.Forms.TextBox txtWO;
        private System.Windows.Forms.Button btnCheck;
        private System.Windows.Forms.Button btnFind;
        private System.Windows.Forms.Label lblOKWorkOrder;
        private System.Windows.Forms.Label lblGroupID;
        private System.Windows.Forms.Label lblQty;
        private System.Windows.Forms.Label lblRev;
        private System.Windows.Forms.Label lblMBPN;
        private System.Windows.Forms.Label lblModel;
        private System.Windows.Forms.Label lblWO;
        private System.Windows.Forms.Label lblLine;
        private System.Windows.Forms.Label lblEndDate;
        private System.Windows.Forms.Label lblBeginDate;
        private System.Windows.Forms.RadioButton rbtnGroup;
        private System.Windows.Forms.RadioButton rbtnRelease;
        private System.Windows.Forms.GroupBox gbDispatchQty;
        private System.Windows.Forms.DataGridView dgvDispatch;
    }
}