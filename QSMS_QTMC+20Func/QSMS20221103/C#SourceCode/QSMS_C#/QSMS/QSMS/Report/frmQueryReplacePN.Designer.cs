namespace QSMS.QSMS.Report
{
    partial class frmQueryReplacePN
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
            this.txtModel = new System.Windows.Forms.TextBox();
            this.txtCompPN = new System.Windows.Forms.TextBox();
            this.btnExcel = new System.Windows.Forms.Button();
            this.btnQuery = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.CompPN = new System.Windows.Forms.Label();
            this.dgSAP_BOM = new System.Windows.Forms.DataGridView();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgSAP_BOM)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.txtModel);
            this.groupBox1.Controls.Add(this.txtCompPN);
            this.groupBox1.Controls.Add(this.btnExcel);
            this.groupBox1.Controls.Add(this.btnQuery);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.CompPN);
            this.groupBox1.Location = new System.Drawing.Point(9, 23);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Size = new System.Drawing.Size(683, 70);
            this.groupBox1.TabIndex = 3;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "SAP_BOM";
            // 
            // txtModel
            // 
            this.txtModel.Location = new System.Drawing.Point(339, 28);
            this.txtModel.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtModel.Name = "txtModel";
            this.txtModel.Size = new System.Drawing.Size(146, 21);
            this.txtModel.TabIndex = 5;
            this.txtModel.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtModel_KeyPress);
            // 
            // txtCompPN
            // 
            this.txtCompPN.Location = new System.Drawing.Point(79, 28);
            this.txtCompPN.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtCompPN.Name = "txtCompPN";
            this.txtCompPN.Size = new System.Drawing.Size(149, 21);
            this.txtCompPN.TabIndex = 4;
            this.txtCompPN.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCompPN_KeyPress);
            // 
            // btnExcel
            // 
            this.btnExcel.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnExcel.Location = new System.Drawing.Point(584, 28);
            this.btnExcel.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(80, 26);
            this.btnExcel.TabIndex = 3;
            this.btnExcel.Text = "Excel";
            this.btnExcel.UseVisualStyleBackColor = false;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // btnQuery
            // 
            this.btnQuery.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnQuery.Location = new System.Drawing.Point(502, 28);
            this.btnQuery.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnQuery.Name = "btnQuery";
            this.btnQuery.Size = new System.Drawing.Size(78, 26);
            this.btnQuery.TabIndex = 2;
            this.btnQuery.Text = "Query";
            this.btnQuery.UseVisualStyleBackColor = false;
            this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.Info;
            this.label2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.ForeColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label2.Location = new System.Drawing.Point(260, 28);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(75, 21);
            this.label2.TabIndex = 1;
            this.label2.Text = "Model";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // CompPN
            // 
            this.CompPN.BackColor = System.Drawing.SystemColors.Info;
            this.CompPN.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.CompPN.Location = new System.Drawing.Point(11, 27);
            this.CompPN.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.CompPN.Name = "CompPN";
            this.CompPN.Size = new System.Drawing.Size(63, 21);
            this.CompPN.TabIndex = 0;
            this.CompPN.Text = "CompPN";
            this.CompPN.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // dgSAP_BOM
            // 
            this.dgSAP_BOM.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgSAP_BOM.Location = new System.Drawing.Point(9, 97);
            this.dgSAP_BOM.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.dgSAP_BOM.Name = "dgSAP_BOM";
            this.dgSAP_BOM.RowTemplate.Height = 27;
            this.dgSAP_BOM.Size = new System.Drawing.Size(683, 306);
            this.dgSAP_BOM.TabIndex = 2;
            // 
            // frmQueryReplacePN
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(714, 422);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.dgSAP_BOM);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "frmQueryReplacePN";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmQueryReplacePN";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmQueryReplacePN_FormClosed);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgSAP_BOM)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox txtModel;
        private System.Windows.Forms.TextBox txtCompPN;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.Button btnQuery;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label CompPN;
        private System.Windows.Forms.DataGridView dgSAP_BOM;
    }
}