namespace QSMS.QSMS.PD
{
    partial class frmUpdateRealQty
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
            this.DID = new System.Windows.Forms.Label();
            this.totalQty = new System.Windows.Forms.Label();
            this.realQty = new System.Windows.Forms.Label();
            this.updateTo = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtDID = new System.Windows.Forms.TextBox();
            this.txtTotalQty = new System.Windows.Forms.TextBox();
            this.txtRealQty = new System.Windows.Forms.TextBox();
            this.txtUpdateTo = new System.Windows.Forms.TextBox();
            this.txtReason = new System.Windows.Forms.TextBox();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnReset = new System.Windows.Forms.Button();
            this.txtMsg = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // DID
            // 
            this.DID.BackColor = System.Drawing.SystemColors.Info;
            this.DID.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.DID.Location = new System.Drawing.Point(15, 32);
            this.DID.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.DID.Name = "DID";
            this.DID.Size = new System.Drawing.Size(46, 23);
            this.DID.TabIndex = 0;
            this.DID.Text = "DID";
            // 
            // totalQty
            // 
            this.totalQty.BackColor = System.Drawing.SystemColors.Info;
            this.totalQty.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.totalQty.Location = new System.Drawing.Point(259, 32);
            this.totalQty.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.totalQty.Name = "totalQty";
            this.totalQty.Size = new System.Drawing.Size(106, 23);
            this.totalQty.TabIndex = 2;
            this.totalQty.Text = "TotalQty";
            // 
            // realQty
            // 
            this.realQty.BackColor = System.Drawing.SystemColors.Info;
            this.realQty.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.realQty.Location = new System.Drawing.Point(430, 32);
            this.realQty.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.realQty.Name = "realQty";
            this.realQty.Size = new System.Drawing.Size(106, 23);
            this.realQty.TabIndex = 4;
            this.realQty.Text = "Real Qty";
            // 
            // updateTo
            // 
            this.updateTo.BackColor = System.Drawing.SystemColors.Info;
            this.updateTo.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.updateTo.Location = new System.Drawing.Point(15, 80);
            this.updateTo.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.updateTo.Name = "updateTo";
            this.updateTo.Size = new System.Drawing.Size(106, 23);
            this.updateTo.TabIndex = 6;
            this.updateTo.Text = "UpdateTo";
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(234, 80);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(82, 23);
            this.label1.TabIndex = 8;
            this.label1.Text = "Reason";
            // 
            // txtDID
            // 
            this.txtDID.Font = new System.Drawing.Font("宋体", 10F);
            this.txtDID.Location = new System.Drawing.Point(62, 32);
            this.txtDID.Margin = new System.Windows.Forms.Padding(2);
            this.txtDID.Name = "txtDID";
            this.txtDID.Size = new System.Drawing.Size(182, 23);
            this.txtDID.TabIndex = 1;
            this.txtDID.TextChanged += new System.EventHandler(this.txtDID_TextChanged);
            this.txtDID.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDID_KeyDown);
            // 
            // txtTotalQty
            // 
            this.txtTotalQty.Font = new System.Drawing.Font("宋体", 10F);
            this.txtTotalQty.Location = new System.Drawing.Point(364, 32);
            this.txtTotalQty.Margin = new System.Windows.Forms.Padding(2);
            this.txtTotalQty.Name = "txtTotalQty";
            this.txtTotalQty.Size = new System.Drawing.Size(62, 23);
            this.txtTotalQty.TabIndex = 3;
            // 
            // txtRealQty
            // 
            this.txtRealQty.Font = new System.Drawing.Font("宋体", 10F);
            this.txtRealQty.Location = new System.Drawing.Point(536, 32);
            this.txtRealQty.Margin = new System.Windows.Forms.Padding(2);
            this.txtRealQty.Name = "txtRealQty";
            this.txtRealQty.Size = new System.Drawing.Size(66, 23);
            this.txtRealQty.TabIndex = 5;
            // 
            // txtUpdateTo
            // 
            this.txtUpdateTo.Font = new System.Drawing.Font("宋体", 10F);
            this.txtUpdateTo.Location = new System.Drawing.Point(121, 80);
            this.txtUpdateTo.Margin = new System.Windows.Forms.Padding(2);
            this.txtUpdateTo.Name = "txtUpdateTo";
            this.txtUpdateTo.Size = new System.Drawing.Size(74, 23);
            this.txtUpdateTo.TabIndex = 7;
            this.txtUpdateTo.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtUpdateTo_KeyDown);
            // 
            // txtReason
            // 
            this.txtReason.Location = new System.Drawing.Point(320, 80);
            this.txtReason.Margin = new System.Windows.Forms.Padding(2);
            this.txtReason.Name = "txtReason";
            this.txtReason.Size = new System.Drawing.Size(282, 21);
            this.txtReason.TabIndex = 9;
            this.txtReason.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtReason_KeyDown);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(13, 138);
            this.richTextBox1.Margin = new System.Windows.Forms.Padding(2);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.ReadOnly = true;
            this.richTextBox1.Size = new System.Drawing.Size(355, 81);
            this.richTextBox1.TabIndex = 10;
            this.richTextBox1.Text = "说明:\n  1.刷入你想要更改数量的DID号码, 会带出它当前的Real Qty。\n  2.在Update To 后面输入你要改的DID数量。\n  3.在Reas" +
    "on后面输入原因，点Save保存。";
            // 
            // btnSave
            // 
            this.btnSave.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnSave.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnSave.Location = new System.Drawing.Point(385, 156);
            this.btnSave.Margin = new System.Windows.Forms.Padding(2);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(100, 34);
            this.btnSave.TabIndex = 11;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = false;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnReset
            // 
            this.btnReset.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnReset.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnReset.Location = new System.Drawing.Point(506, 157);
            this.btnReset.Margin = new System.Windows.Forms.Padding(2);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(108, 32);
            this.btnReset.TabIndex = 12;
            this.btnReset.Text = "Reset";
            this.btnReset.UseVisualStyleBackColor = false;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // txtMsg
            // 
            this.txtMsg.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.txtMsg.Location = new System.Drawing.Point(13, 234);
            this.txtMsg.Margin = new System.Windows.Forms.Padding(2);
            this.txtMsg.Name = "txtMsg";
            this.txtMsg.ReadOnly = true;
            this.txtMsg.Size = new System.Drawing.Size(603, 21);
            this.txtMsg.TabIndex = 13;
            // 
            // frmUpdateRealQty
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(701, 272);
            this.Controls.Add(this.txtMsg);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.txtReason);
            this.Controls.Add(this.txtUpdateTo);
            this.Controls.Add(this.txtRealQty);
            this.Controls.Add(this.txtTotalQty);
            this.Controls.Add(this.txtDID);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.updateTo);
            this.Controls.Add(this.realQty);
            this.Controls.Add(this.totalQty);
            this.Controls.Add(this.DID);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "frmUpdateRealQty";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmUpdateReelQty";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmUpdateRealQty_FormClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label DID;
        private System.Windows.Forms.Label totalQty;
        private System.Windows.Forms.Label realQty;
        private System.Windows.Forms.Label updateTo;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtDID;
        private System.Windows.Forms.TextBox txtTotalQty;
        private System.Windows.Forms.TextBox txtRealQty;
        private System.Windows.Forms.TextBox txtUpdateTo;
        private System.Windows.Forms.TextBox txtReason;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.TextBox txtMsg;
    }
}