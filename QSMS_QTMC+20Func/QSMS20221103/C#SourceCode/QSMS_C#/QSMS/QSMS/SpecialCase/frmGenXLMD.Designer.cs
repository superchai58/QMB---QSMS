namespace QSMS.QSMS.SpecialCase
{
    partial class frmGenXLMD
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
            this.cobFac = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cobType = new System.Windows.Forms.ComboBox();
            this.btnGenXLMD = new System.Windows.Forms.Button();
            this.txtmsg = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(24, 43);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "Factory";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cobFac
            // 
            this.cobFac.FormattingEnabled = true;
            this.cobFac.Location = new System.Drawing.Point(111, 43);
            this.cobFac.Margin = new System.Windows.Forms.Padding(2);
            this.cobFac.Name = "cobFac";
            this.cobFac.Size = new System.Drawing.Size(141, 20);
            this.cobFac.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.Info;
            this.label2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(267, 43);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 20);
            this.label2.TabIndex = 2;
            this.label2.Text = "Type";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // cobType
            // 
            this.cobType.FormattingEnabled = true;
            this.cobType.Location = new System.Drawing.Point(336, 43);
            this.cobType.Margin = new System.Windows.Forms.Padding(2);
            this.cobType.Name = "cobType";
            this.cobType.Size = new System.Drawing.Size(141, 20);
            this.cobType.TabIndex = 3;
            // 
            // btnGenXLMD
            // 
            this.btnGenXLMD.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnGenXLMD.Font = new System.Drawing.Font("宋体", 10F);
            this.btnGenXLMD.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btnGenXLMD.Location = new System.Drawing.Point(501, 37);
            this.btnGenXLMD.Margin = new System.Windows.Forms.Padding(2);
            this.btnGenXLMD.Name = "btnGenXLMD";
            this.btnGenXLMD.Size = new System.Drawing.Size(92, 34);
            this.btnGenXLMD.TabIndex = 4;
            this.btnGenXLMD.Text = "GenXLMD";
            this.btnGenXLMD.UseVisualStyleBackColor = false;
            this.btnGenXLMD.Click += new System.EventHandler(this.btnGenXLMD_Click);
            // 
            // txtmsg
            // 
            this.txtmsg.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.txtmsg.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtmsg.ForeColor = System.Drawing.Color.Red;
            this.txtmsg.Location = new System.Drawing.Point(27, 93);
            this.txtmsg.Margin = new System.Windows.Forms.Padding(2);
            this.txtmsg.Multiline = true;
            this.txtmsg.Name = "txtmsg";
            this.txtmsg.Size = new System.Drawing.Size(566, 102);
            this.txtmsg.TabIndex = 5;
            this.txtmsg.Text = "注意:\r\n      可以再次计算XL需求的时间是第一次XL跑过1H~5H之间\r\n例如:\r\n      XL时间为7:40 那么可以再次计算需求的时间段为8:40" +
    "~12:40,\r\n如果超过这个时间点将不允许需手动跑,将在由系统自动计算.";
            // 
            // frmGenXLMD
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(642, 214);
            this.Controls.Add(this.txtmsg);
            this.Controls.Add(this.btnGenXLMD);
            this.Controls.Add(this.cobType);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cobFac);
            this.Controls.Add(this.label1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "frmGenXLMD";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmGenXLMD";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmGenXLMD_FormClosed);
            this.Load += new System.EventHandler(this.frmGenXLMD_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cobFac;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cobType;
        private System.Windows.Forms.Button btnGenXLMD;
        private System.Windows.Forms.TextBox txtmsg;
    }
}