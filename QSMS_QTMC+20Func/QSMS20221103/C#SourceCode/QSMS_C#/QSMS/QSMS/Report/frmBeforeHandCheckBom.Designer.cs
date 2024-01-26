namespace QSMS.QSMS.Report
{
    partial class frmBeforeHandCheckBom
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
            this.btnExcel = new System.Windows.Forms.Button();
            this.btnCheckBom = new System.Windows.Forms.Button();
            this.dgResult = new System.Windows.Forms.DataGridView();
            this.cboFactory = new System.Windows.Forms.ComboBox();
            this.cboLine = new System.Windows.Forms.ComboBox();
            this.cboModel = new System.Windows.Forms.ComboBox();
            this.txtCQty = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgResult)).BeginInit();
            this.SuspendLayout();
            // 
            // btnExcel
            // 
            this.btnExcel.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnExcel.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnExcel.Location = new System.Drawing.Point(611, 87);
            this.btnExcel.Margin = new System.Windows.Forms.Padding(2);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(79, 30);
            this.btnExcel.TabIndex = 24;
            this.btnExcel.Text = "Excel";
            this.btnExcel.UseVisualStyleBackColor = false;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // btnCheckBom
            // 
            this.btnCheckBom.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnCheckBom.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnCheckBom.Location = new System.Drawing.Point(611, 44);
            this.btnCheckBom.Margin = new System.Windows.Forms.Padding(2);
            this.btnCheckBom.Name = "btnCheckBom";
            this.btnCheckBom.Size = new System.Drawing.Size(79, 30);
            this.btnCheckBom.TabIndex = 23;
            this.btnCheckBom.Text = "CheckBom";
            this.btnCheckBom.UseVisualStyleBackColor = false;
            this.btnCheckBom.Click += new System.EventHandler(this.btnCheckBom_Click);
            // 
            // dgResult
            // 
            this.dgResult.BackgroundColor = System.Drawing.Color.White;
            this.dgResult.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgResult.Location = new System.Drawing.Point(11, 121);
            this.dgResult.Margin = new System.Windows.Forms.Padding(2);
            this.dgResult.Name = "dgResult";
            this.dgResult.RowTemplate.Height = 27;
            this.dgResult.Size = new System.Drawing.Size(718, 291);
            this.dgResult.TabIndex = 22;
            // 
            // cboFactory
            // 
            this.cboFactory.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cboFactory.FormattingEnabled = true;
            this.cboFactory.Location = new System.Drawing.Point(137, 85);
            this.cboFactory.Margin = new System.Windows.Forms.Padding(2);
            this.cboFactory.Name = "cboFactory";
            this.cboFactory.Size = new System.Drawing.Size(144, 22);
            this.cboFactory.TabIndex = 21;
            // 
            // cboLine
            // 
            this.cboLine.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cboLine.FormattingEnabled = true;
            this.cboLine.Location = new System.Drawing.Point(407, 40);
            this.cboLine.Margin = new System.Windows.Forms.Padding(2);
            this.cboLine.Name = "cboLine";
            this.cboLine.Size = new System.Drawing.Size(144, 22);
            this.cboLine.TabIndex = 20;
            // 
            // cboModel
            // 
            this.cboModel.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cboModel.FormattingEnabled = true;
            this.cboModel.Location = new System.Drawing.Point(142, 42);
            this.cboModel.Margin = new System.Windows.Forms.Padding(2);
            this.cboModel.Name = "cboModel";
            this.cboModel.Size = new System.Drawing.Size(144, 22);
            this.cboModel.TabIndex = 19;
            // 
            // txtCQty
            // 
            this.txtCQty.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtCQty.Location = new System.Drawing.Point(407, 87);
            this.txtCQty.Margin = new System.Windows.Forms.Padding(2);
            this.txtCQty.Name = "txtCQty";
            this.txtCQty.Size = new System.Drawing.Size(144, 23);
            this.txtCQty.TabIndex = 18;
            this.txtCQty.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCQty_KeyPress);
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.SystemColors.Info;
            this.label4.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(315, 85);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(88, 23);
            this.label4.TabIndex = 17;
            this.label4.Text = "CombineQty";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.SystemColors.Info;
            this.label3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(315, 40);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(88, 23);
            this.label3.TabIndex = 16;
            this.label3.Text = "Line";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.Info;
            this.label2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(51, 85);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(82, 23);
            this.label2.TabIndex = 15;
            this.label2.Text = "Factory";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(56, 41);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(82, 23);
            this.label1.TabIndex = 14;
            this.label1.Text = "ModelName";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // frmBeforeHandCheckBom
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(735, 428);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.btnCheckBom);
            this.Controls.Add(this.dgResult);
            this.Controls.Add(this.cboFactory);
            this.Controls.Add(this.cboLine);
            this.Controls.Add(this.cboModel);
            this.Controls.Add(this.txtCQty);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "frmBeforeHandCheckBom";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmBeforeHandCheckBom";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmBeforeHandCheckBom_FormClosed);
            this.Load += new System.EventHandler(this.frmBeforeHandCheckBom_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgResult)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.Button btnCheckBom;
        private System.Windows.Forms.DataGridView dgResult;
        private System.Windows.Forms.ComboBox cboFactory;
        private System.Windows.Forms.ComboBox cboLine;
        private System.Windows.Forms.ComboBox cboModel;
        private System.Windows.Forms.TextBox txtCQty;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
    }
}