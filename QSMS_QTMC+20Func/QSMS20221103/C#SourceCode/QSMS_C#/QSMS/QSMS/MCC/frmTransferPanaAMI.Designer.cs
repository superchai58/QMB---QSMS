namespace QSMS.QSMS.MCC
{
    partial class frmTransferPanaAMI
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.cmdGetMEBom = new System.Windows.Forms.Button();
            this.cmdSelect = new System.Windows.Forms.Button();
            this.txtFile = new System.Windows.Forms.TextBox();
            this.chkNewVersion = new System.Windows.Forms.CheckBox();
            this.chkAutChkBom = new System.Windows.Forms.CheckBox();
            this.LabelRun = new System.Windows.Forms.Label();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabel3 = new System.Windows.Forms.ToolStripStatusLabel();
            this.panel1.SuspendLayout();
            this.menuStrip1.SuspendLayout();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.cmdGetMEBom);
            this.panel1.Controls.Add(this.cmdSelect);
            this.panel1.Controls.Add(this.txtFile);
            this.panel1.Controls.Add(this.chkNewVersion);
            this.panel1.Controls.Add(this.chkAutChkBom);
            this.panel1.Location = new System.Drawing.Point(12, 35);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(950, 75);
            this.panel1.TabIndex = 0;
            // 
            // cmdGetMEBom
            // 
            this.cmdGetMEBom.BackColor = System.Drawing.SystemColors.Info;
            this.cmdGetMEBom.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdGetMEBom.Location = new System.Drawing.Point(861, 12);
            this.cmdGetMEBom.Name = "cmdGetMEBom";
            this.cmdGetMEBom.Size = new System.Drawing.Size(70, 33);
            this.cmdGetMEBom.TabIndex = 4;
            this.cmdGetMEBom.Text = "Upload";
            this.cmdGetMEBom.UseVisualStyleBackColor = false;
            this.cmdGetMEBom.Click += new System.EventHandler(this.cmdGetMEBom_Click);
            // 
            // cmdSelect
            // 
            this.cmdSelect.BackColor = System.Drawing.SystemColors.Info;
            this.cmdSelect.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdSelect.Location = new System.Drawing.Point(641, 12);
            this.cmdSelect.Name = "cmdSelect";
            this.cmdSelect.Size = new System.Drawing.Size(87, 34);
            this.cmdSelect.TabIndex = 3;
            this.cmdSelect.Text = "Select";
            this.cmdSelect.UseVisualStyleBackColor = false;
            this.cmdSelect.Click += new System.EventHandler(this.cmdSelect_Click);
            // 
            // txtFile
            // 
            this.txtFile.Location = new System.Drawing.Point(137, 22);
            this.txtFile.Name = "txtFile";
            this.txtFile.Size = new System.Drawing.Size(488, 23);
            this.txtFile.TabIndex = 2;
            // 
            // chkNewVersion
            // 
            this.chkNewVersion.AutoSize = true;
            this.chkNewVersion.BackColor = System.Drawing.SystemColors.Info;
            this.chkNewVersion.Enabled = false;
            this.chkNewVersion.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.chkNewVersion.Location = new System.Drawing.Point(734, 20);
            this.chkNewVersion.Name = "chkNewVersion";
            this.chkNewVersion.Size = new System.Drawing.Size(115, 20);
            this.chkNewVersion.TabIndex = 1;
            this.chkNewVersion.Text = "New Version";
            this.chkNewVersion.UseVisualStyleBackColor = false;
            // 
            // chkAutChkBom
            // 
            this.chkAutChkBom.AutoSize = true;
            this.chkAutChkBom.BackColor = System.Drawing.SystemColors.Info;
            this.chkAutChkBom.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.chkAutChkBom.ForeColor = System.Drawing.Color.Black;
            this.chkAutChkBom.Location = new System.Drawing.Point(12, 23);
            this.chkAutChkBom.Name = "chkAutChkBom";
            this.chkAutChkBom.Size = new System.Drawing.Size(123, 20);
            this.chkAutChkBom.TabIndex = 0;
            this.chkAutChkBom.Text = "AutoCheckBom";
            this.chkAutChkBom.UseVisualStyleBackColor = false;
            // 
            // LabelRun
            // 
            this.LabelRun.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.LabelRun.Font = new System.Drawing.Font("宋体", 10.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.LabelRun.Location = new System.Drawing.Point(8, 129);
            this.LabelRun.Name = "LabelRun";
            this.LabelRun.Size = new System.Drawing.Size(954, 31);
            this.LabelRun.TabIndex = 30;
            // 
            // menuStrip1
            // 
            this.menuStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exitToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(990, 25);
            this.menuStrip1.TabIndex = 31;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(40, 21);
            this.exitToolStripMenuItem.Text = "Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.statusStrip1.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1,
            this.toolStripStatusLabel2,
            this.toolStripStatusLabel3});
            this.statusStrip1.Location = new System.Drawing.Point(0, 171);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(990, 22);
            this.statusStrip1.TabIndex = 32;
            this.statusStrip1.Text = "statusStrip1";
            this.statusStrip1.ItemClicked += new System.Windows.Forms.ToolStripItemClickedEventHandler(this.statusStrip1_ItemClicked);
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(0, 17);
            // 
            // toolStripStatusLabel2
            // 
            this.toolStripStatusLabel2.Name = "toolStripStatusLabel2";
            this.toolStripStatusLabel2.Size = new System.Drawing.Size(0, 17);
            // 
            // toolStripStatusLabel3
            // 
            this.toolStripStatusLabel3.Name = "toolStripStatusLabel3";
            this.toolStripStatusLabel3.Size = new System.Drawing.Size(0, 17);
            // 
            // frmTransferPanaAMI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(990, 193);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.LabelRun);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.menuStrip1);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.MainMenuStrip = this.menuStrip1;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "frmTransferPanaAMI";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "TransPanaMAI[20230406]";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmTransferPanaAMI_FormClosed);
            this.Load += new System.EventHandler(this.frmTransferPanaAMI_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button cmdGetMEBom;
        private System.Windows.Forms.Button cmdSelect;
        private System.Windows.Forms.TextBox txtFile;
        private System.Windows.Forms.CheckBox chkNewVersion;
        private System.Windows.Forms.CheckBox chkAutChkBom;
        private System.Windows.Forms.Label LabelRun;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel2;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel3;
    }
}