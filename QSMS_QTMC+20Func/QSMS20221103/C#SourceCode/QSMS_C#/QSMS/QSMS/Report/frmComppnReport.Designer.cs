namespace QSMS.QSMS.Report
{
    partial class frmComppnReport
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
            this.textWO1 = new System.Windows.Forms.TextBox();
            this.textWO2 = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cboType = new System.Windows.Forms.ComboBox();
            this.button1 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(47, 53);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 23);
            this.label1.TabIndex = 0;
            this.label1.Text = "WO1:";
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.Info;
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(47, 90);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(71, 21);
            this.label2.TabIndex = 1;
            this.label2.Text = "WO2:";
            // 
            // textWO1
            // 
            this.textWO1.Location = new System.Drawing.Point(115, 55);
            this.textWO1.Name = "textWO1";
            this.textWO1.Size = new System.Drawing.Size(183, 21);
            this.textWO1.TabIndex = 1;
            this.textWO1.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textWO1_KeyPress);
            // 
            // textWO2
            // 
            this.textWO2.Location = new System.Drawing.Point(115, 90);
            this.textWO2.Name = "textWO2";
            this.textWO2.Size = new System.Drawing.Size(183, 21);
            this.textWO2.TabIndex = 2;
            this.textWO2.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.textWO2_KeyPress);
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.SystemColors.Info;
            this.label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(47, 122);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(71, 23);
            this.label3.TabIndex = 4;
            this.label3.Text = "Type:";
            // 
            // cboType
            // 
            this.cboType.FormattingEnabled = true;
            this.cboType.Location = new System.Drawing.Point(115, 122);
            this.cboType.Name = "cboType";
            this.cboType.Size = new System.Drawing.Size(183, 20);
            this.cboType.TabIndex = 3;
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.button1.Location = new System.Drawing.Point(340, 75);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(84, 40);
            this.button1.TabIndex = 4;
            this.button1.Text = "EXECL";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // frmComppnReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(493, 207);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.cboType);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textWO2);
            this.Controls.Add(this.textWO1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Name = "frmComppnReport";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmComppnReport";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmComppnReport_FormClosed);
            this.Load += new System.EventHandler(this.frmComppnReport_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textWO1;
        private System.Windows.Forms.TextBox textWO2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cboType;
        private System.Windows.Forms.Button button1;
    }
}