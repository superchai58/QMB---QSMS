namespace QSMS.QSMS.PD
{
    partial class FrmDIDBake
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
            this.txtBakeDID = new System.Windows.Forms.TextBox();
            this.cmdBakeQ = new System.Windows.Forms.Button();
            this.DataGridDIDBake = new System.Windows.Forms.DataGridView();
            this.cmdBakeOK = new System.Windows.Forms.Button();
            this.cmdEndBake = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.DataGridDIDBake)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.Font = new System.Drawing.Font("宋体", 15F);
            this.label1.Location = new System.Drawing.Point(33, 33);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(69, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = " DID：";
            // 
            // txtBakeDID
            // 
            this.txtBakeDID.Font = new System.Drawing.Font("宋体", 12F);
            this.txtBakeDID.Location = new System.Drawing.Point(104, 31);
            this.txtBakeDID.Margin = new System.Windows.Forms.Padding(2);
            this.txtBakeDID.Name = "txtBakeDID";
            this.txtBakeDID.Size = new System.Drawing.Size(303, 26);
            this.txtBakeDID.TabIndex = 1;
            // 
            // cmdBakeQ
            // 
            this.cmdBakeQ.Font = new System.Drawing.Font("宋体", 12F);
            this.cmdBakeQ.Location = new System.Drawing.Point(423, 27);
            this.cmdBakeQ.Margin = new System.Windows.Forms.Padding(2);
            this.cmdBakeQ.Name = "cmdBakeQ";
            this.cmdBakeQ.Size = new System.Drawing.Size(75, 30);
            this.cmdBakeQ.TabIndex = 2;
            this.cmdBakeQ.Text = "Query";
            this.cmdBakeQ.UseVisualStyleBackColor = true;
            this.cmdBakeQ.Click += new System.EventHandler(this.cmdBakeQ_Click);
            // 
            // DataGridDIDBake
            // 
            this.DataGridDIDBake.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.DataGridDIDBake.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DataGridDIDBake.Location = new System.Drawing.Point(11, 67);
            this.DataGridDIDBake.Margin = new System.Windows.Forms.Padding(2);
            this.DataGridDIDBake.Name = "DataGridDIDBake";
            this.DataGridDIDBake.RowTemplate.Height = 37;
            this.DataGridDIDBake.Size = new System.Drawing.Size(693, 216);
            this.DataGridDIDBake.TabIndex = 5;
            // 
            // cmdBakeOK
            // 
            this.cmdBakeOK.Font = new System.Drawing.Font("宋体", 12F);
            this.cmdBakeOK.Location = new System.Drawing.Point(511, 27);
            this.cmdBakeOK.Margin = new System.Windows.Forms.Padding(2);
            this.cmdBakeOK.Name = "cmdBakeOK";
            this.cmdBakeOK.Size = new System.Drawing.Size(86, 30);
            this.cmdBakeOK.TabIndex = 6;
            this.cmdBakeOK.Text = "StartBake";
            this.cmdBakeOK.UseVisualStyleBackColor = true;
            this.cmdBakeOK.Click += new System.EventHandler(this.cmdBakeOK_Click);
            // 
            // cmdEndBake
            // 
            this.cmdEndBake.Font = new System.Drawing.Font("宋体", 12F);
            this.cmdEndBake.Location = new System.Drawing.Point(601, 27);
            this.cmdEndBake.Margin = new System.Windows.Forms.Padding(2);
            this.cmdEndBake.Name = "cmdEndBake";
            this.cmdEndBake.Size = new System.Drawing.Size(75, 30);
            this.cmdEndBake.TabIndex = 7;
            this.cmdEndBake.Text = "EndBake";
            this.cmdEndBake.UseVisualStyleBackColor = true;
            this.cmdEndBake.Click += new System.EventHandler(this.cmdEndBake_Click);
            // 
            // FrmDIDBake
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(705, 294);
            this.Controls.Add(this.cmdEndBake);
            this.Controls.Add(this.cmdBakeOK);
            this.Controls.Add(this.DataGridDIDBake);
            this.Controls.Add(this.cmdBakeQ);
            this.Controls.Add(this.txtBakeDID);
            this.Controls.Add(this.label1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "FrmDIDBake";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FrmDIDBake";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FrmDIDBake_FormClosed);
            this.Load += new System.EventHandler(this.FrmDIDBake_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DataGridDIDBake)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtBakeDID;
        private System.Windows.Forms.Button cmdBakeQ;
        private System.Windows.Forms.DataGridView DataGridDIDBake;
        private System.Windows.Forms.Button cmdBakeOK;
        private System.Windows.Forms.Button cmdEndBake;
    }
}