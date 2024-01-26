namespace QSMS.QSMS.MCC
{
    partial class frmQryReturnedDID
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
            this.lblDID = new System.Windows.Forms.Label();
            this.lblDIDA = new System.Windows.Forms.Label();
            this.txtDID = new System.Windows.Forms.TextBox();
            this.txtNewDID = new System.Windows.Forms.TextBox();
            this.btnQuery = new System.Windows.Forms.Button();
            this.btnQueryA = new System.Windows.Forms.Button();
            this.btnExcel = new System.Windows.Forms.Button();
            this.dgReturnedDID = new System.Windows.Forms.DataGridView();
            ((System.ComponentModel.ISupportInitialize)(this.dgReturnedDID)).BeginInit();
            this.SuspendLayout();
            // 
            // lblDID
            // 
            this.lblDID.BackColor = System.Drawing.SystemColors.Info;
            this.lblDID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblDID.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblDID.Location = new System.Drawing.Point(44, 18);
            this.lblDID.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblDID.Name = "lblDID";
            this.lblDID.Size = new System.Drawing.Size(75, 29);
            this.lblDID.TabIndex = 0;
            this.lblDID.Text = "DID";
            this.lblDID.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblDIDA
            // 
            this.lblDIDA.BackColor = System.Drawing.SystemColors.Info;
            this.lblDIDA.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblDIDA.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblDIDA.Location = new System.Drawing.Point(44, 60);
            this.lblDIDA.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblDIDA.Name = "lblDIDA";
            this.lblDIDA.Size = new System.Drawing.Size(75, 31);
            this.lblDIDA.TabIndex = 1;
            this.lblDIDA.Text = "DID-A";
            this.lblDIDA.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtDID
            // 
            this.txtDID.Location = new System.Drawing.Point(120, 18);
            this.txtDID.Margin = new System.Windows.Forms.Padding(2);
            this.txtDID.Multiline = true;
            this.txtDID.Name = "txtDID";
            this.txtDID.Size = new System.Drawing.Size(329, 29);
            this.txtDID.TabIndex = 2;
            this.txtDID.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDID_KeyPress);
            // 
            // txtNewDID
            // 
            this.txtNewDID.Location = new System.Drawing.Point(120, 60);
            this.txtNewDID.Margin = new System.Windows.Forms.Padding(2);
            this.txtNewDID.Multiline = true;
            this.txtNewDID.Name = "txtNewDID";
            this.txtNewDID.Size = new System.Drawing.Size(329, 31);
            this.txtNewDID.TabIndex = 3;
            this.txtNewDID.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtNewDID_KeyPress);
            // 
            // btnQuery
            // 
            this.btnQuery.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnQuery.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnQuery.Location = new System.Drawing.Point(472, 18);
            this.btnQuery.Margin = new System.Windows.Forms.Padding(2);
            this.btnQuery.Name = "btnQuery";
            this.btnQuery.Size = new System.Drawing.Size(76, 29);
            this.btnQuery.TabIndex = 4;
            this.btnQuery.Text = "Query";
            this.btnQuery.UseVisualStyleBackColor = false;
            this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // btnQueryA
            // 
            this.btnQueryA.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnQueryA.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnQueryA.Location = new System.Drawing.Point(472, 60);
            this.btnQueryA.Margin = new System.Windows.Forms.Padding(2);
            this.btnQueryA.Name = "btnQueryA";
            this.btnQueryA.Size = new System.Drawing.Size(76, 31);
            this.btnQueryA.TabIndex = 5;
            this.btnQueryA.Text = "Query-A";
            this.btnQueryA.UseVisualStyleBackColor = false;
            this.btnQueryA.Click += new System.EventHandler(this.btnQueryA_Click);
            // 
            // btnExcel
            // 
            this.btnExcel.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnExcel.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnExcel.Location = new System.Drawing.Point(561, 18);
            this.btnExcel.Margin = new System.Windows.Forms.Padding(2);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(74, 29);
            this.btnExcel.TabIndex = 6;
            this.btnExcel.Text = "Excel";
            this.btnExcel.UseVisualStyleBackColor = false;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // dgReturnedDID
            // 
            this.dgReturnedDID.BackgroundColor = System.Drawing.SystemColors.ControlLightLight;
            this.dgReturnedDID.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgReturnedDID.Location = new System.Drawing.Point(11, 107);
            this.dgReturnedDID.Margin = new System.Windows.Forms.Padding(2);
            this.dgReturnedDID.Name = "dgReturnedDID";
            this.dgReturnedDID.RowTemplate.Height = 27;
            this.dgReturnedDID.Size = new System.Drawing.Size(639, 352);
            this.dgReturnedDID.TabIndex = 7;
            // 
            // frmQryReturnedDID
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(665, 470);
            this.Controls.Add(this.dgReturnedDID);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.btnQueryA);
            this.Controls.Add(this.btnQuery);
            this.Controls.Add(this.txtNewDID);
            this.Controls.Add(this.txtDID);
            this.Controls.Add(this.lblDIDA);
            this.Controls.Add(this.lblDID);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "frmQryReturnedDID";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmQryReturnedDID";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmQryReturnedDID_FormClosed);
            ((System.ComponentModel.ISupportInitialize)(this.dgReturnedDID)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblDID;
        private System.Windows.Forms.Label lblDIDA;
        private System.Windows.Forms.TextBox txtDID;
        private System.Windows.Forms.TextBox txtNewDID;
        private System.Windows.Forms.Button btnQuery;
        private System.Windows.Forms.Button btnQueryA;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.DataGridView dgReturnedDID;
    }
}