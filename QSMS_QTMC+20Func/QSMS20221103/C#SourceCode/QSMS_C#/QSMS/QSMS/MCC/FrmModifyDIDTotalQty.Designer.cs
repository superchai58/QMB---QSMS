namespace QSMS.QSMS.MCC
{
    partial class FrmModifyDIDTotalQty
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
            this.DG1 = new System.Windows.Forms.DataGridView();
            this.cmdExit = new System.Windows.Forms.Button();
            this.CmdRefresh = new System.Windows.Forms.Button();
            this.cmdCancel = new System.Windows.Forms.Button();
            this.cmdSave = new System.Windows.Forms.Button();
            this.cmdUpdate = new System.Windows.Forms.Button();
            this.cmdFind = new System.Windows.Forms.Button();
            this.CboVendorCode = new System.Windows.Forms.ComboBox();
            this.CboLotCode = new System.Windows.Forms.ComboBox();
            this.CboDID = new System.Windows.Forms.ComboBox();
            this.CboDateCode = new System.Windows.Forms.ComboBox();
            this.CboCompPN = new System.Windows.Forms.ComboBox();
            this.TxtQty = new System.Windows.Forms.TextBox();
            this.TxtGroupQty = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            ((System.ComponentModel.ISupportInitialize)(this.DG1)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // DG1
            // 
            this.DG1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DG1.Location = new System.Drawing.Point(5, 157);
            this.DG1.Margin = new System.Windows.Forms.Padding(4);
            this.DG1.Name = "DG1";
            this.DG1.RowTemplate.Height = 27;
            this.DG1.Size = new System.Drawing.Size(1071, 443);
            this.DG1.TabIndex = 22;
            this.DG1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DG1_CellClick);
            this.DG1.SelectionChanged += new System.EventHandler(this.DG1_SelectionChanged);
            // 
            // cmdExit
            // 
            this.cmdExit.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdExit.Location = new System.Drawing.Point(754, 93);
            this.cmdExit.Margin = new System.Windows.Forms.Padding(4);
            this.cmdExit.Name = "cmdExit";
            this.cmdExit.Size = new System.Drawing.Size(94, 40);
            this.cmdExit.TabIndex = 19;
            this.cmdExit.Text = "EXIT";
            this.cmdExit.UseVisualStyleBackColor = true;
            this.cmdExit.Click += new System.EventHandler(this.cmdExit_Click);
            // 
            // CmdRefresh
            // 
            this.CmdRefresh.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CmdRefresh.Location = new System.Drawing.Point(648, 93);
            this.CmdRefresh.Margin = new System.Windows.Forms.Padding(4);
            this.CmdRefresh.Name = "CmdRefresh";
            this.CmdRefresh.Size = new System.Drawing.Size(94, 40);
            this.CmdRefresh.TabIndex = 18;
            this.CmdRefresh.Text = "Refresh";
            this.CmdRefresh.UseVisualStyleBackColor = true;
            this.CmdRefresh.Click += new System.EventHandler(this.CmdRefresh_Click);
            // 
            // cmdCancel
            // 
            this.cmdCancel.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdCancel.Location = new System.Drawing.Point(540, 94);
            this.cmdCancel.Margin = new System.Windows.Forms.Padding(4);
            this.cmdCancel.Name = "cmdCancel";
            this.cmdCancel.Size = new System.Drawing.Size(94, 40);
            this.cmdCancel.TabIndex = 17;
            this.cmdCancel.Text = "&Cancel";
            this.cmdCancel.UseVisualStyleBackColor = true;
            this.cmdCancel.Click += new System.EventHandler(this.cmdCancel_Click);
            // 
            // cmdSave
            // 
            this.cmdSave.Enabled = false;
            this.cmdSave.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdSave.Location = new System.Drawing.Point(432, 93);
            this.cmdSave.Margin = new System.Windows.Forms.Padding(4);
            this.cmdSave.Name = "cmdSave";
            this.cmdSave.Size = new System.Drawing.Size(94, 40);
            this.cmdSave.TabIndex = 16;
            this.cmdSave.Text = "&Save";
            this.cmdSave.UseVisualStyleBackColor = true;
            this.cmdSave.Click += new System.EventHandler(this.cmdSave_Click);
            // 
            // cmdUpdate
            // 
            this.cmdUpdate.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdUpdate.Location = new System.Drawing.Point(326, 93);
            this.cmdUpdate.Margin = new System.Windows.Forms.Padding(4);
            this.cmdUpdate.Name = "cmdUpdate";
            this.cmdUpdate.Size = new System.Drawing.Size(94, 40);
            this.cmdUpdate.TabIndex = 15;
            this.cmdUpdate.Text = "&Update";
            this.cmdUpdate.UseVisualStyleBackColor = true;
            this.cmdUpdate.Click += new System.EventHandler(this.cmdUpdate_Click);
            // 
            // cmdFind
            // 
            this.cmdFind.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdFind.Location = new System.Drawing.Point(220, 93);
            this.cmdFind.Margin = new System.Windows.Forms.Padding(4);
            this.cmdFind.Name = "cmdFind";
            this.cmdFind.Size = new System.Drawing.Size(94, 40);
            this.cmdFind.TabIndex = 14;
            this.cmdFind.Text = "&Find";
            this.cmdFind.UseVisualStyleBackColor = true;
            this.cmdFind.Click += new System.EventHandler(this.cmdFind_Click);
            // 
            // CboVendorCode
            // 
            this.CboVendorCode.FormattingEnabled = true;
            this.CboVendorCode.Location = new System.Drawing.Point(620, 17);
            this.CboVendorCode.Margin = new System.Windows.Forms.Padding(4);
            this.CboVendorCode.Name = "CboVendorCode";
            this.CboVendorCode.Size = new System.Drawing.Size(148, 25);
            this.CboVendorCode.TabIndex = 13;
            this.CboVendorCode.Click += new System.EventHandler(this.CboVendorCode_Click);
            this.CboVendorCode.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.CboVendorCode_KeyPress);
            // 
            // CboLotCode
            // 
            this.CboLotCode.FormattingEnabled = true;
            this.CboLotCode.Location = new System.Drawing.Point(122, 51);
            this.CboLotCode.Margin = new System.Windows.Forms.Padding(4);
            this.CboLotCode.Name = "CboLotCode";
            this.CboLotCode.Size = new System.Drawing.Size(120, 25);
            this.CboLotCode.TabIndex = 12;
            this.CboLotCode.Click += new System.EventHandler(this.CboLotCode_Click);
            this.CboLotCode.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.CboLotCode_KeyPress);
            // 
            // CboDID
            // 
            this.CboDID.FormattingEnabled = true;
            this.CboDID.Location = new System.Drawing.Point(620, 51);
            this.CboDID.Margin = new System.Windows.Forms.Padding(4);
            this.CboDID.Name = "CboDID";
            this.CboDID.Size = new System.Drawing.Size(411, 25);
            this.CboDID.TabIndex = 11;
            this.CboDID.Click += new System.EventHandler(this.CboDID_Click);
            // 
            // CboDateCode
            // 
            this.CboDateCode.FormattingEnabled = true;
            this.CboDateCode.Location = new System.Drawing.Point(888, 16);
            this.CboDateCode.Margin = new System.Windows.Forms.Padding(4);
            this.CboDateCode.Name = "CboDateCode";
            this.CboDateCode.Size = new System.Drawing.Size(143, 25);
            this.CboDateCode.TabIndex = 10;
            this.CboDateCode.Click += new System.EventHandler(this.CboDateCode_Click);
            this.CboDateCode.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.CboDateCode_KeyPress);
            // 
            // CboCompPN
            // 
            this.CboCompPN.FormattingEnabled = true;
            this.CboCompPN.Location = new System.Drawing.Point(366, 17);
            this.CboCompPN.Margin = new System.Windows.Forms.Padding(4);
            this.CboCompPN.Name = "CboCompPN";
            this.CboCompPN.Size = new System.Drawing.Size(143, 25);
            this.CboCompPN.TabIndex = 9;
            this.CboCompPN.Click += new System.EventHandler(this.CboCompPN_Click);
            this.CboCompPN.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.CboCompPN_KeyPress);
            // 
            // TxtQty
            // 
            this.TxtQty.Location = new System.Drawing.Point(367, 53);
            this.TxtQty.Margin = new System.Windows.Forms.Padding(4);
            this.TxtQty.Name = "TxtQty";
            this.TxtQty.Size = new System.Drawing.Size(143, 23);
            this.TxtQty.TabIndex = 8;
            this.TxtQty.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.TxtQty_KeyPress);
            // 
            // TxtGroupQty
            // 
            this.TxtGroupQty.Location = new System.Drawing.Point(122, 19);
            this.TxtGroupQty.Margin = new System.Windows.Forms.Padding(4);
            this.TxtGroupQty.Name = "TxtGroupQty";
            this.TxtGroupQty.Size = new System.Drawing.Size(120, 23);
            this.TxtGroupQty.TabIndex = 7;
            this.TxtGroupQty.Text = "1";
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.SystemColors.Info;
            this.label7.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label7.Location = new System.Drawing.Point(517, 53);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(100, 23);
            this.label7.TabIndex = 6;
            this.label7.Text = "DID";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.SystemColors.Info;
            this.label6.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.Location = new System.Drawing.Point(19, 53);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(100, 23);
            this.label6.TabIndex = 5;
            this.label6.Text = "Lot Code";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.SystemColors.Info;
            this.label5.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.Location = new System.Drawing.Point(517, 19);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(100, 23);
            this.label5.TabIndex = 4;
            this.label5.Text = "Vendor Code";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.SystemColors.Info;
            this.label4.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(250, 53);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(113, 23);
            this.label4.TabIndex = 3;
            this.label4.Text = "Qty";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.SystemColors.Info;
            this.label3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(776, 18);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(109, 23);
            this.label3.TabIndex = 2;
            this.label3.Text = "Date Code";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.Info;
            this.label2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(250, 19);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(113, 23);
            this.label2.TabIndex = 1;
            this.label2.Text = "CompPN";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(19, 18);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(100, 23);
            this.label1.TabIndex = 0;
            this.label1.Text = "Group Qty";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.cmdExit);
            this.panel1.Controls.Add(this.CmdRefresh);
            this.panel1.Controls.Add(this.cmdCancel);
            this.panel1.Controls.Add(this.cmdSave);
            this.panel1.Controls.Add(this.cmdUpdate);
            this.panel1.Controls.Add(this.cmdFind);
            this.panel1.Controls.Add(this.CboVendorCode);
            this.panel1.Controls.Add(this.CboLotCode);
            this.panel1.Controls.Add(this.CboDID);
            this.panel1.Controls.Add(this.CboDateCode);
            this.panel1.Controls.Add(this.CboCompPN);
            this.panel1.Controls.Add(this.TxtQty);
            this.panel1.Controls.Add(this.TxtGroupQty);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(5, 4);
            this.panel1.Margin = new System.Windows.Forms.Padding(5);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1071, 144);
            this.panel1.TabIndex = 21;
            // 
            // FrmModifyDIDTotalQty
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1081, 613);
            this.Controls.Add(this.DG1);
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "FrmModifyDIDTotalQty";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FrmModifyDIDTotalQty";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FrmModifyDIDTotalQty_FormClosed);
            this.Load += new System.EventHandler(this.FrmModifyDIDTotalQty_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DG1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView DG1;
        private System.Windows.Forms.Button cmdExit;
        private System.Windows.Forms.Button CmdRefresh;
        private System.Windows.Forms.Button cmdCancel;
        private System.Windows.Forms.Button cmdSave;
        private System.Windows.Forms.Button cmdUpdate;
        private System.Windows.Forms.Button cmdFind;
        private System.Windows.Forms.ComboBox CboVendorCode;
        private System.Windows.Forms.ComboBox CboLotCode;
        private System.Windows.Forms.ComboBox CboDID;
        private System.Windows.Forms.ComboBox CboDateCode;
        private System.Windows.Forms.ComboBox CboCompPN;
        private System.Windows.Forms.TextBox TxtQty;
        private System.Windows.Forms.TextBox TxtGroupQty;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panel1;
    }
}