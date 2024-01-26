namespace QSMS.QSMS.MCC
{
    partial class frmInput
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
            this.lblmsg = new System.Windows.Forms.Label();
            this.txtInput = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // lblmsg
            // 
            this.lblmsg.BackColor = System.Drawing.SystemColors.Info;
            this.lblmsg.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblmsg.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblmsg.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.lblmsg.Location = new System.Drawing.Point(38, 26);
            this.lblmsg.Name = "lblmsg";
            this.lblmsg.Size = new System.Drawing.Size(90, 28);
            this.lblmsg.TabIndex = 94;
            this.lblmsg.Text = "请输入:";
            this.lblmsg.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtInput
            // 
            this.txtInput.Font = new System.Drawing.Font("微软雅黑", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtInput.Location = new System.Drawing.Point(134, 25);
            this.txtInput.Name = "txtInput";
            this.txtInput.Size = new System.Drawing.Size(339, 29);
            this.txtInput.TabIndex = 93;
            this.txtInput.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtInput_KeyDown);
            // 
            // frmInput
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(520, 83);
            this.Controls.Add(this.lblmsg);
            this.Controls.Add(this.txtInput);
            this.MaximizeBox = false;
            this.Name = "frmInput";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "输入框";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblmsg;
        private System.Windows.Forms.TextBox txtInput;
    }
}