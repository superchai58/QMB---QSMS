namespace QSMS.QSMS.MCC
{
    partial class IC_Burn
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
            this.txtDID = new System.Windows.Forms.TextBox();
            this.txtCompPN = new System.Windows.Forms.TextBox();
            this.txtModelName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.cboPN = new System.Windows.Forms.ComboBox();
            this.btnLinkshearpin = new System.Windows.Forms.Button();
            this.btnccshearpin = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtDID
            // 
            this.txtDID.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtDID.Location = new System.Drawing.Point(44, 39);
            this.txtDID.Name = "txtDID";
            this.txtDID.Size = new System.Drawing.Size(193, 23);
            this.txtDID.TabIndex = 0;
            this.txtDID.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDID_KeyPress);
            // 
            // txtCompPN
            // 
            this.txtCompPN.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtCompPN.Location = new System.Drawing.Point(336, 42);
            this.txtCompPN.Name = "txtCompPN";
            this.txtCompPN.Size = new System.Drawing.Size(212, 23);
            this.txtCompPN.TabIndex = 1;
            // 
            // txtModelName
            // 
            this.txtModelName.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtModelName.Location = new System.Drawing.Point(334, 119);
            this.txtModelName.Name = "txtModelName";
            this.txtModelName.Size = new System.Drawing.Size(212, 23);
            this.txtModelName.TabIndex = 3;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.label1.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(15, 42);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(28, 14);
            this.label1.TabIndex = 4;
            this.label1.Text = "DID";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.label2.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(271, 45);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(49, 14);
            this.label2.TabIndex = 5;
            this.label2.Text = "CompPN";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.label3.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(17, 122);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(21, 14);
            this.label3.TabIndex = 6;
            this.label3.Text = "PN";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.label4.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(260, 122);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(70, 14);
            this.label4.TabIndex = 7;
            this.label4.Text = "ModelName";
            // 
            // cboPN
            // 
            this.cboPN.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cboPN.FormattingEnabled = true;
            this.cboPN.Location = new System.Drawing.Point(44, 119);
            this.cboPN.Name = "cboPN";
            this.cboPN.Size = new System.Drawing.Size(193, 22);
            this.cboPN.TabIndex = 8;
            this.cboPN.SelectedIndexChanged += new System.EventHandler(this.cboPN_SelectedIndexChanged);
            // 
            // btnLinkshearpin
            // 
            this.btnLinkshearpin.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnLinkshearpin.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnLinkshearpin.Location = new System.Drawing.Point(65, 197);
            this.btnLinkshearpin.Name = "btnLinkshearpin";
            this.btnLinkshearpin.Size = new System.Drawing.Size(122, 23);
            this.btnLinkshearpin.TabIndex = 9;
            this.btnLinkshearpin.Text = "DIDLinkShearPin";
            this.btnLinkshearpin.UseVisualStyleBackColor = false;
            this.btnLinkshearpin.Click += new System.EventHandler(this.btnLinkshearpin_Click);
            // 
            // btnccshearpin
            // 
            this.btnccshearpin.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnccshearpin.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnccshearpin.Location = new System.Drawing.Point(356, 197);
            this.btnccshearpin.Name = "btnccshearpin";
            this.btnccshearpin.Size = new System.Drawing.Size(142, 23);
            this.btnccshearpin.TabIndex = 10;
            this.btnccshearpin.Text = "Query_CCShearPin";
            this.btnccshearpin.UseVisualStyleBackColor = false;
            this.btnccshearpin.Click += new System.EventHandler(this.btnccshearpin_Click);
            // 
            // IC_Burn
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(558, 276);
            this.Controls.Add(this.btnccshearpin);
            this.Controls.Add(this.btnLinkshearpin);
            this.Controls.Add(this.cboPN);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtModelName);
            this.Controls.Add(this.txtCompPN);
            this.Controls.Add(this.txtDID);
            this.Name = "IC_Burn";
            this.Text = "IC_Burn";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtDID;
        private System.Windows.Forms.TextBox txtCompPN;
        private System.Windows.Forms.TextBox txtModelName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cboPN;
        private System.Windows.Forms.Button btnLinkshearpin;
        private System.Windows.Forms.Button btnccshearpin;
    }
}