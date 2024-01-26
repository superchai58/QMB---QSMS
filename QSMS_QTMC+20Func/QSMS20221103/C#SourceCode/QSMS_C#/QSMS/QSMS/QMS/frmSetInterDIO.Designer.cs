namespace QSMS.QSMS.QMS
{
    partial class frmSetInterDIO
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
            this.CboLine = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.chkAllLine = new System.Windows.Forms.CheckBox();
            this.cmdOK = new System.Windows.Forms.Button();
            this.cmdExit = new System.Windows.Forms.Button();
            this.fraMachine = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox1.Controls.Add(this.CboLine);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.chkAllLine);
            this.groupBox1.Location = new System.Drawing.Point(2, 15);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox1.Size = new System.Drawing.Size(659, 71);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "by Line or Machine";
            // 
            // CboLine
            // 
            this.CboLine.FormattingEnabled = true;
            this.CboLine.Location = new System.Drawing.Point(484, 28);
            this.CboLine.Margin = new System.Windows.Forms.Padding(4);
            this.CboLine.Name = "CboLine";
            this.CboLine.Size = new System.Drawing.Size(150, 25);
            this.CboLine.TabIndex = 4;
            this.CboLine.SelectedValueChanged += new System.EventHandler(this.CboLine_SelectedValueChanged);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.ForeColor = System.Drawing.Color.Black;
            this.label1.Location = new System.Drawing.Point(359, 28);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(117, 25);
            this.label1.TabIndex = 3;
            this.label1.Text = "Set Machine";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // chkAllLine
            // 
            this.chkAllLine.BackColor = System.Drawing.SystemColors.Info;
            this.chkAllLine.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.chkAllLine.ForeColor = System.Drawing.Color.Black;
            this.chkAllLine.Location = new System.Drawing.Point(31, 28);
            this.chkAllLine.Margin = new System.Windows.Forms.Padding(4);
            this.chkAllLine.Name = "chkAllLine";
            this.chkAllLine.Size = new System.Drawing.Size(108, 21);
            this.chkAllLine.TabIndex = 2;
            this.chkAllLine.Text = "SET ALL LINE";
            this.chkAllLine.UseVisualStyleBackColor = false;
            this.chkAllLine.Click += new System.EventHandler(this.chkAllLine_Click);
            // 
            // cmdOK
            // 
            this.cmdOK.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.cmdOK.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdOK.Location = new System.Drawing.Point(669, 30);
            this.cmdOK.Margin = new System.Windows.Forms.Padding(4);
            this.cmdOK.Name = "cmdOK";
            this.cmdOK.Size = new System.Drawing.Size(75, 34);
            this.cmdOK.TabIndex = 0;
            this.cmdOK.Text = "OK";
            this.cmdOK.UseVisualStyleBackColor = false;
            this.cmdOK.Click += new System.EventHandler(this.cmdOK_Click);
            // 
            // cmdExit
            // 
            this.cmdExit.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.cmdExit.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdExit.Location = new System.Drawing.Point(771, 32);
            this.cmdExit.Margin = new System.Windows.Forms.Padding(4);
            this.cmdExit.Name = "cmdExit";
            this.cmdExit.Size = new System.Drawing.Size(73, 34);
            this.cmdExit.TabIndex = 1;
            this.cmdExit.Text = "Exit";
            this.cmdExit.UseVisualStyleBackColor = false;
            this.cmdExit.Click += new System.EventHandler(this.cmdExit_Click);
            // 
            // fraMachine
            // 
            this.fraMachine.BackColor = System.Drawing.SystemColors.Control;
            this.fraMachine.Location = new System.Drawing.Point(2, 93);
            this.fraMachine.Name = "fraMachine";
            this.fraMachine.Size = new System.Drawing.Size(842, 251);
            this.fraMachine.TabIndex = 2;
            this.fraMachine.TabStop = false;
            this.fraMachine.Text = "Machine in Line";
            // 
            // frmSetInterDIO
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(869, 358);
            this.Controls.Add(this.fraMachine);
            this.Controls.Add(this.cmdExit);
            this.Controls.Add(this.cmdOK);
            this.Controls.Add(this.groupBox1);
            this.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "frmSetInterDIO";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Set DIO and InterLock[20100810]";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmSetInterDIO_FormClosed);
            this.Load += new System.EventHandler(this.frmSetInterDIO_Load);
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox CboLine;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chkAllLine;
        private System.Windows.Forms.Button cmdOK;
        private System.Windows.Forms.Button cmdExit;
        private System.Windows.Forms.GroupBox fraMachine;
    }
}