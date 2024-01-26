namespace QSMS.QSMS
{
    partial class frmPrinterSetting
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
            this.gbPrinter = new System.Windows.Forms.GroupBox();
            this.rbtnSATO = new System.Windows.Forms.RadioButton();
            this.rbtnZebra = new System.Windows.Forms.RadioButton();
            this.gbPort = new System.Windows.Forms.GroupBox();
            this.rbtnNetwork = new System.Windows.Forms.RadioButton();
            this.rbtnLPT = new System.Windows.Forms.RadioButton();
            this.rbtnCom = new System.Windows.Forms.RadioButton();
            this.gbdpm = new System.Windows.Forms.GroupBox();
            this.rbtn200 = new System.Windows.Forms.RadioButton();
            this.rbtn300 = new System.Windows.Forms.RadioButton();
            this.lblCompPort = new System.Windows.Forms.Label();
            this.lblSettings = new System.Windows.Forms.Label();
            this.txtCompPort = new System.Windows.Forms.TextBox();
            this.txtport1 = new System.Windows.Forms.TextBox();
            this.txtport2 = new System.Windows.Forms.TextBox();
            this.txtport3 = new System.Windows.Forms.TextBox();
            this.txtport4 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.gbPrinter.SuspendLayout();
            this.gbPort.SuspendLayout();
            this.gbdpm.SuspendLayout();
            this.SuspendLayout();
            // 
            // gbPrinter
            // 
            this.gbPrinter.Controls.Add(this.rbtnSATO);
            this.gbPrinter.Controls.Add(this.rbtnZebra);
            this.gbPrinter.Location = new System.Drawing.Point(12, 13);
            this.gbPrinter.Name = "gbPrinter";
            this.gbPrinter.Size = new System.Drawing.Size(449, 53);
            this.gbPrinter.TabIndex = 0;
            this.gbPrinter.TabStop = false;
            this.gbPrinter.Text = "Printer";
            // 
            // rbtnSATO
            // 
            this.rbtnSATO.AutoSize = true;
            this.rbtnSATO.BackColor = System.Drawing.SystemColors.Info;
            this.rbtnSATO.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rbtnSATO.Location = new System.Drawing.Point(173, 20);
            this.rbtnSATO.Name = "rbtnSATO";
            this.rbtnSATO.Size = new System.Drawing.Size(122, 20);
            this.rbtnSATO.TabIndex = 1;
            this.rbtnSATO.Text = "SATO printer";
            this.rbtnSATO.UseVisualStyleBackColor = false;
            // 
            // rbtnZebra
            // 
            this.rbtnZebra.AutoSize = true;
            this.rbtnZebra.BackColor = System.Drawing.SystemColors.Info;
            this.rbtnZebra.Checked = true;
            this.rbtnZebra.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rbtnZebra.Location = new System.Drawing.Point(13, 20);
            this.rbtnZebra.Name = "rbtnZebra";
            this.rbtnZebra.Size = new System.Drawing.Size(130, 20);
            this.rbtnZebra.TabIndex = 0;
            this.rbtnZebra.TabStop = true;
            this.rbtnZebra.Text = "Zebra printer";
            this.rbtnZebra.UseVisualStyleBackColor = false;
            // 
            // gbPort
            // 
            this.gbPort.Controls.Add(this.rbtnNetwork);
            this.gbPort.Controls.Add(this.rbtnLPT);
            this.gbPort.Controls.Add(this.rbtnCom);
            this.gbPort.Location = new System.Drawing.Point(12, 72);
            this.gbPort.Name = "gbPort";
            this.gbPort.Size = new System.Drawing.Size(449, 52);
            this.gbPort.TabIndex = 1;
            this.gbPort.TabStop = false;
            this.gbPort.Text = "Port";
            // 
            // rbtnNetwork
            // 
            this.rbtnNetwork.AutoSize = true;
            this.rbtnNetwork.BackColor = System.Drawing.SystemColors.Info;
            this.rbtnNetwork.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rbtnNetwork.Location = new System.Drawing.Point(315, 21);
            this.rbtnNetwork.Name = "rbtnNetwork";
            this.rbtnNetwork.Size = new System.Drawing.Size(82, 20);
            this.rbtnNetwork.TabIndex = 2;
            this.rbtnNetwork.Text = "Network";
            this.rbtnNetwork.UseVisualStyleBackColor = false;
            // 
            // rbtnLPT
            // 
            this.rbtnLPT.AutoSize = true;
            this.rbtnLPT.BackColor = System.Drawing.SystemColors.Info;
            this.rbtnLPT.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rbtnLPT.Location = new System.Drawing.Point(173, 21);
            this.rbtnLPT.Name = "rbtnLPT";
            this.rbtnLPT.Size = new System.Drawing.Size(90, 20);
            this.rbtnLPT.TabIndex = 1;
            this.rbtnLPT.Text = "LPT Port";
            this.rbtnLPT.UseVisualStyleBackColor = false;
            // 
            // rbtnCom
            // 
            this.rbtnCom.AutoSize = true;
            this.rbtnCom.BackColor = System.Drawing.SystemColors.Info;
            this.rbtnCom.Checked = true;
            this.rbtnCom.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rbtnCom.Location = new System.Drawing.Point(13, 21);
            this.rbtnCom.Name = "rbtnCom";
            this.rbtnCom.Size = new System.Drawing.Size(98, 20);
            this.rbtnCom.TabIndex = 0;
            this.rbtnCom.TabStop = true;
            this.rbtnCom.Text = "Comp Port";
            this.rbtnCom.UseVisualStyleBackColor = false;
            // 
            // gbdpm
            // 
            this.gbdpm.Controls.Add(this.rbtn200);
            this.gbdpm.Controls.Add(this.rbtn300);
            this.gbdpm.Location = new System.Drawing.Point(12, 141);
            this.gbdpm.Name = "gbdpm";
            this.gbdpm.Size = new System.Drawing.Size(449, 50);
            this.gbdpm.TabIndex = 2;
            this.gbdpm.TabStop = false;
            this.gbdpm.Text = "dpm";
            // 
            // rbtn200
            // 
            this.rbtn200.AutoSize = true;
            this.rbtn200.BackColor = System.Drawing.SystemColors.Info;
            this.rbtn200.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rbtn200.Location = new System.Drawing.Point(173, 20);
            this.rbtn200.Name = "rbtn200";
            this.rbtn200.Size = new System.Drawing.Size(50, 20);
            this.rbtn200.TabIndex = 1;
            this.rbtn200.Text = "200";
            this.rbtn200.UseVisualStyleBackColor = false;
            // 
            // rbtn300
            // 
            this.rbtn300.AutoSize = true;
            this.rbtn300.BackColor = System.Drawing.SystemColors.Info;
            this.rbtn300.Checked = true;
            this.rbtn300.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.rbtn300.Location = new System.Drawing.Point(13, 20);
            this.rbtn300.Name = "rbtn300";
            this.rbtn300.Size = new System.Drawing.Size(50, 20);
            this.rbtn300.TabIndex = 0;
            this.rbtn300.TabStop = true;
            this.rbtn300.Text = "300";
            this.rbtn300.UseVisualStyleBackColor = false;
            // 
            // lblCompPort
            // 
            this.lblCompPort.BackColor = System.Drawing.SystemColors.Info;
            this.lblCompPort.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblCompPort.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblCompPort.Location = new System.Drawing.Point(12, 216);
            this.lblCompPort.Name = "lblCompPort";
            this.lblCompPort.Size = new System.Drawing.Size(90, 21);
            this.lblCompPort.TabIndex = 3;
            this.lblCompPort.Text = "CompPort:";
            this.lblCompPort.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // lblSettings
            // 
            this.lblSettings.BackColor = System.Drawing.SystemColors.Info;
            this.lblSettings.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblSettings.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblSettings.Location = new System.Drawing.Point(12, 258);
            this.lblSettings.Name = "lblSettings";
            this.lblSettings.Size = new System.Drawing.Size(90, 21);
            this.lblSettings.TabIndex = 4;
            this.lblSettings.Text = "Settings:";
            this.lblSettings.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtCompPort
            // 
            this.txtCompPort.Location = new System.Drawing.Point(108, 216);
            this.txtCompPort.Name = "txtCompPort";
            this.txtCompPort.Size = new System.Drawing.Size(118, 21);
            this.txtCompPort.TabIndex = 5;
            this.txtCompPort.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // txtport1
            // 
            this.txtport1.Location = new System.Drawing.Point(108, 258);
            this.txtport1.Name = "txtport1";
            this.txtport1.Size = new System.Drawing.Size(89, 21);
            this.txtport1.TabIndex = 6;
            this.txtport1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // txtport2
            // 
            this.txtport2.Location = new System.Drawing.Point(214, 258);
            this.txtport2.Name = "txtport2";
            this.txtport2.Size = new System.Drawing.Size(57, 21);
            this.txtport2.TabIndex = 7;
            this.txtport2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // txtport3
            // 
            this.txtport3.Location = new System.Drawing.Point(288, 258);
            this.txtport3.Name = "txtport3";
            this.txtport3.Size = new System.Drawing.Size(45, 21);
            this.txtport3.TabIndex = 8;
            this.txtport3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // txtport4
            // 
            this.txtport4.Location = new System.Drawing.Point(350, 258);
            this.txtport4.Name = "txtport4";
            this.txtport4.Size = new System.Drawing.Size(50, 21);
            this.txtport4.TabIndex = 9;
            this.txtport4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.label1.Location = new System.Drawing.Point(197, 258);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(17, 21);
            this.label1.TabIndex = 10;
            this.label1.Text = "--";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.label2.Location = new System.Drawing.Point(271, 258);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(17, 21);
            this.label2.TabIndex = 11;
            this.label2.Text = "--";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.label3.Location = new System.Drawing.Point(333, 258);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(17, 21);
            this.label3.TabIndex = 12;
            this.label3.Text = "--";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(66, 296);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 13;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnExit
            // 
            this.btnExit.Location = new System.Drawing.Point(239, 296);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 23);
            this.btnExit.TabIndex = 14;
            this.btnExit.Text = "Exit";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // frmPrinterSetting
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(480, 342);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtport4);
            this.Controls.Add(this.txtport3);
            this.Controls.Add(this.txtport2);
            this.Controls.Add(this.txtport1);
            this.Controls.Add(this.txtCompPort);
            this.Controls.Add(this.lblSettings);
            this.Controls.Add(this.lblCompPort);
            this.Controls.Add(this.gbdpm);
            this.Controls.Add(this.gbPort);
            this.Controls.Add(this.gbPrinter);
            this.Name = "frmPrinterSetting";
            this.Text = "frmPrinterSetting  [V20221031]";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmPrinterSetting_FormClosed);
            this.Load += new System.EventHandler(this.frmPrinterSetting_Load);
            this.gbPrinter.ResumeLayout(false);
            this.gbPrinter.PerformLayout();
            this.gbPort.ResumeLayout(false);
            this.gbPort.PerformLayout();
            this.gbdpm.ResumeLayout(false);
            this.gbdpm.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox gbPrinter;
        private System.Windows.Forms.RadioButton rbtnSATO;
        private System.Windows.Forms.RadioButton rbtnZebra;
        private System.Windows.Forms.GroupBox gbPort;
        private System.Windows.Forms.RadioButton rbtnNetwork;
        private System.Windows.Forms.RadioButton rbtnLPT;
        private System.Windows.Forms.RadioButton rbtnCom;
        private System.Windows.Forms.GroupBox gbdpm;
        private System.Windows.Forms.RadioButton rbtn200;
        private System.Windows.Forms.RadioButton rbtn300;
        private System.Windows.Forms.Label lblCompPort;
        private System.Windows.Forms.Label lblSettings;
        private System.Windows.Forms.TextBox txtCompPort;
        private System.Windows.Forms.TextBox txtport1;
        private System.Windows.Forms.TextBox txtport2;
        private System.Windows.Forms.TextBox txtport3;
        private System.Windows.Forms.TextBox txtport4;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnExit;
    }
}