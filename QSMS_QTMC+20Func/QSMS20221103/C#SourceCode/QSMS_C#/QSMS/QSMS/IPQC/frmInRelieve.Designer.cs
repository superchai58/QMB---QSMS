namespace QSMS.QSMS.IPQC
{
    partial class frmInRelieve
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
            this.btnRelieve = new System.Windows.Forms.Button();
            this.lblDID = new System.Windows.Forms.Label();
            this.txtDID = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnRelieve
            // 
            this.btnRelieve.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnRelieve.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnRelieve.Location = new System.Drawing.Point(377, 36);
            this.btnRelieve.Name = "btnRelieve";
            this.btnRelieve.Size = new System.Drawing.Size(87, 30);
            this.btnRelieve.TabIndex = 0;
            this.btnRelieve.Text = "Relieve";
            this.btnRelieve.UseVisualStyleBackColor = false;
            this.btnRelieve.Click += new System.EventHandler(this.btnRelieve_Click);
            // 
            // lblDID
            // 
            this.lblDID.BackColor = System.Drawing.SystemColors.Info;
            this.lblDID.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblDID.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblDID.Location = new System.Drawing.Point(29, 42);
            this.lblDID.Name = "lblDID";
            this.lblDID.Size = new System.Drawing.Size(65, 21);
            this.lblDID.TabIndex = 1;
            this.lblDID.Text = "DID:";
            this.lblDID.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtDID
            // 
            this.txtDID.Location = new System.Drawing.Point(100, 42);
            this.txtDID.Name = "txtDID";
            this.txtDID.Size = new System.Drawing.Size(236, 21);
            this.txtDID.TabIndex = 2;
            // 
            // frmInRelieve
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(543, 115);
            this.Controls.Add(this.txtDID);
            this.Controls.Add(this.lblDID);
            this.Controls.Add(this.btnRelieve);
            this.Name = "frmInRelieve";
            this.Text = "RelieveIPQC";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmInRelieve_FormClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnRelieve;
        private System.Windows.Forms.Label lblDID;
        private System.Windows.Forms.TextBox txtDID;
    }
}