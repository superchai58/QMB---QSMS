namespace QSMS.QSMS.SpecialCase
{
    partial class frmUpdateUID
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
            this.DGUidinfo = new System.Windows.Forms.DataGridView();
            this.btnQuery = new System.Windows.Forms.Button();
            this.btnUpdate = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtOldUID = new System.Windows.Forms.TextBox();
            this.txtNewUID = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.DGUidinfo)).BeginInit();
            this.SuspendLayout();
            // 
            // DGUidinfo
            // 
            this.DGUidinfo.BackgroundColor = System.Drawing.Color.White;
            this.DGUidinfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DGUidinfo.Location = new System.Drawing.Point(27, 59);
            this.DGUidinfo.Margin = new System.Windows.Forms.Padding(2);
            this.DGUidinfo.Name = "DGUidinfo";
            this.DGUidinfo.RowTemplate.Height = 27;
            this.DGUidinfo.Size = new System.Drawing.Size(670, 232);
            this.DGUidinfo.TabIndex = 0;
            this.DGUidinfo.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DGUidinfo_CellContentClick);
            // 
            // btnQuery
            // 
            this.btnQuery.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnQuery.Location = new System.Drawing.Point(505, 22);
            this.btnQuery.Margin = new System.Windows.Forms.Padding(2);
            this.btnQuery.Name = "btnQuery";
            this.btnQuery.Size = new System.Drawing.Size(90, 27);
            this.btnQuery.TabIndex = 1;
            this.btnQuery.Text = "查询";
            this.btnQuery.UseVisualStyleBackColor = false;
            this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // btnUpdate
            // 
            this.btnUpdate.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnUpdate.Location = new System.Drawing.Point(610, 22);
            this.btnUpdate.Margin = new System.Windows.Forms.Padding(2);
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.Size = new System.Drawing.Size(87, 27);
            this.btnUpdate.TabIndex = 2;
            this.btnUpdate.Text = "更新";
            this.btnUpdate.UseVisualStyleBackColor = false;
            this.btnUpdate.Click += new System.EventHandler(this.btnUpdate_Click);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(24, 22);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 21);
            this.label1.TabIndex = 3;
            this.label1.Text = "旧工号:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.Info;
            this.label2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(265, 22);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(75, 21);
            this.label2.TabIndex = 4;
            this.label2.Text = "新工号：";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // txtOldUID
            // 
            this.txtOldUID.Location = new System.Drawing.Point(111, 22);
            this.txtOldUID.Margin = new System.Windows.Forms.Padding(2);
            this.txtOldUID.Name = "txtOldUID";
            this.txtOldUID.Size = new System.Drawing.Size(140, 21);
            this.txtOldUID.TabIndex = 5;
            // 
            // txtNewUID
            // 
            this.txtNewUID.Location = new System.Drawing.Point(352, 22);
            this.txtNewUID.Margin = new System.Windows.Forms.Padding(2);
            this.txtNewUID.Name = "txtNewUID";
            this.txtNewUID.Size = new System.Drawing.Size(140, 21);
            this.txtNewUID.TabIndex = 6;
            // 
            // frmUpdateUID
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(737, 317);
            this.Controls.Add(this.txtNewUID);
            this.Controls.Add(this.txtOldUID);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnUpdate);
            this.Controls.Add(this.btnQuery);
            this.Controls.Add(this.DGUidinfo);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "frmUpdateUID";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmUpdateUID";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmUpdateUID_FormClosed);
            ((System.ComponentModel.ISupportInitialize)(this.DGUidinfo)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView DGUidinfo;
        private System.Windows.Forms.Button btnQuery;
        private System.Windows.Forms.Button btnUpdate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtOldUID;
        private System.Windows.Forms.TextBox txtNewUID;
    }
}