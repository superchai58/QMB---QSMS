namespace QSMS.QSMS.QMS
{
    partial class frmSendXLRemainDemand
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btnSend = new System.Windows.Forms.Button();
            this.btnQuery = new System.Windows.Forms.Button();
            this.dgResult1 = new System.Windows.Forms.DataGridView();
            this.label3 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtDate = new System.Windows.Forms.TextBox();
            this.txtShift = new System.Windows.Forms.TextBox();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.dgResult2 = new System.Windows.Forms.DataGridView();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.dgResult3 = new System.Windows.Forms.DataGridView();
            this.btnExcel = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.txtFactory = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dgResult1)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgResult2)).BeginInit();
            this.tabPage3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgResult3)).BeginInit();
            this.SuspendLayout();
            // 
            // btnSend
            // 
            this.btnSend.BackColor = System.Drawing.Color.Yellow;
            this.btnSend.Font = new System.Drawing.Font("SimSun", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnSend.Location = new System.Drawing.Point(1098, 47);
            this.btnSend.Margin = new System.Windows.Forms.Padding(2);
            this.btnSend.Name = "btnSend";
            this.btnSend.Size = new System.Drawing.Size(79, 33);
            this.btnSend.TabIndex = 24;
            this.btnSend.Text = "Send";
            this.btnSend.UseVisualStyleBackColor = false;
            this.btnSend.Click += new System.EventHandler(this.btnSend_Click);
            // 
            // btnQuery
            // 
            this.btnQuery.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnQuery.Font = new System.Drawing.Font("SimSun", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnQuery.Location = new System.Drawing.Point(739, 43);
            this.btnQuery.Margin = new System.Windows.Forms.Padding(2);
            this.btnQuery.Name = "btnQuery";
            this.btnQuery.Size = new System.Drawing.Size(79, 33);
            this.btnQuery.TabIndex = 23;
            this.btnQuery.Text = "Query";
            this.btnQuery.UseVisualStyleBackColor = false;
            this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // dgResult1
            // 
            this.dgResult1.BackgroundColor = System.Drawing.Color.White;
            this.dgResult1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgResult1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgResult1.Location = new System.Drawing.Point(3, 3);
            this.dgResult1.Margin = new System.Windows.Forms.Padding(2);
            this.dgResult1.Name = "dgResult1";
            this.dgResult1.RowTemplate.Height = 20;
            this.dgResult1.Size = new System.Drawing.Size(1192, 595);
            this.dgResult1.TabIndex = 22;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.SystemColors.Info;
            this.label3.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(269, 43);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(82, 25);
            this.label3.TabIndex = 16;
            this.label3.Text = "Shift";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(20, 43);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(82, 25);
            this.label1.TabIndex = 14;
            this.label1.Text = "Date";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // txtDate
            // 
            this.txtDate.Font = new System.Drawing.Font("SimSun", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtDate.Location = new System.Drawing.Point(106, 43);
            this.txtDate.Margin = new System.Windows.Forms.Padding(2);
            this.txtDate.Name = "txtDate";
            this.txtDate.Size = new System.Drawing.Size(144, 23);
            this.txtDate.TabIndex = 25;
            // 
            // txtShift
            // 
            this.txtShift.Font = new System.Drawing.Font("SimSun", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtShift.Location = new System.Drawing.Point(355, 43);
            this.txtShift.Margin = new System.Windows.Forms.Padding(2);
            this.txtShift.Name = "txtShift";
            this.txtShift.Size = new System.Drawing.Size(113, 23);
            this.txtShift.TabIndex = 26;
            // 
            // tabControl1
            // 
            this.tabControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Controls.Add(this.tabPage3);
            this.tabControl1.Location = new System.Drawing.Point(23, 102);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(1206, 627);
            this.tabControl1.TabIndex = 27;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.dgResult1);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(1198, 601);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Result1";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.dgResult2);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(1198, 601);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Result2";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // dgResult2
            // 
            this.dgResult2.BackgroundColor = System.Drawing.Color.White;
            this.dgResult2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgResult2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgResult2.Location = new System.Drawing.Point(3, 3);
            this.dgResult2.Margin = new System.Windows.Forms.Padding(2);
            this.dgResult2.Name = "dgResult2";
            this.dgResult2.RowTemplate.Height = 20;
            this.dgResult2.Size = new System.Drawing.Size(1192, 595);
            this.dgResult2.TabIndex = 28;
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.dgResult3);
            this.tabPage3.Location = new System.Drawing.Point(4, 22);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(1198, 601);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Result3";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // dgResult3
            // 
            this.dgResult3.BackgroundColor = System.Drawing.Color.White;
            this.dgResult3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgResult3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgResult3.Location = new System.Drawing.Point(3, 3);
            this.dgResult3.Margin = new System.Windows.Forms.Padding(2);
            this.dgResult3.Name = "dgResult3";
            this.dgResult3.RowTemplate.Height = 20;
            this.dgResult3.Size = new System.Drawing.Size(1192, 595);
            this.dgResult3.TabIndex = 28;
            // 
            // btnExcel
            // 
            this.btnExcel.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnExcel.Font = new System.Drawing.Font("SimSun", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnExcel.Location = new System.Drawing.Point(837, 43);
            this.btnExcel.Margin = new System.Windows.Forms.Padding(2);
            this.btnExcel.Name = "btnExcel";
            this.btnExcel.Size = new System.Drawing.Size(79, 33);
            this.btnExcel.TabIndex = 28;
            this.btnExcel.Text = "导出Excel";
            this.btnExcel.UseVisualStyleBackColor = false;
            this.btnExcel.Click += new System.EventHandler(this.btnExcel_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.SystemColors.Info;
            this.label2.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(490, 47);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(64, 16);
            this.label2.TabIndex = 29;
            this.label2.Text = "Factory";
            // 
            // txtFactory
            // 
            this.txtFactory.Font = new System.Drawing.Font("SimSun", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtFactory.Location = new System.Drawing.Point(560, 43);
            this.txtFactory.Name = "txtFactory";
            this.txtFactory.Size = new System.Drawing.Size(127, 26);
            this.txtFactory.TabIndex = 30;
            // 
            // frmSendXLRemainDemand
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1241, 732);
            this.Controls.Add(this.txtFactory);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnExcel);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.txtShift);
            this.Controls.Add(this.txtDate);
            this.Controls.Add(this.btnSend);
            this.Controls.Add(this.btnQuery);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label1);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "frmSendXLRemainDemand";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmSendXLRemainDemand 20231103";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmSendXLRemainDemand_FormClosed);
            this.Load += new System.EventHandler(this.frmSendXLRemainDemand_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgResult1)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgResult2)).EndInit();
            this.tabPage3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgResult3)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSend;
        private System.Windows.Forms.Button btnQuery;
        private System.Windows.Forms.DataGridView dgResult1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtDate;
        private System.Windows.Forms.TextBox txtShift;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.DataGridView dgResult2;
        private System.Windows.Forms.DataGridView dgResult3;
        private System.Windows.Forms.Button btnExcel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtFactory;
    }
}



