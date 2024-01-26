namespace QSMS.QSMS.SpecialCase
{
    partial class frmUrgentDIDToWH
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.panel1 = new System.Windows.Forms.Panel();
            this.cmdDelete = new System.Windows.Forms.Button();
            this.cmdUpdate = new System.Windows.Forms.Button();
            this.cmdQuery = new System.Windows.Forms.Button();
            this.txtRefID = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dgInfo = new System.Windows.Forms.DataGridView();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgInfo)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.cmdDelete);
            this.panel1.Controls.Add(this.cmdUpdate);
            this.panel1.Controls.Add(this.cmdQuery);
            this.panel1.Controls.Add(this.txtRefID);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(14, 5);
            this.panel1.Margin = new System.Windows.Forms.Padding(2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(644, 54);
            this.panel1.TabIndex = 0;
            // 
            // cmdDelete
            // 
            this.cmdDelete.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.cmdDelete.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdDelete.Location = new System.Drawing.Point(410, 13);
            this.cmdDelete.Margin = new System.Windows.Forms.Padding(2);
            this.cmdDelete.Name = "cmdDelete";
            this.cmdDelete.Size = new System.Drawing.Size(70, 30);
            this.cmdDelete.TabIndex = 4;
            this.cmdDelete.Text = "Delete";
            this.cmdDelete.UseVisualStyleBackColor = false;
            this.cmdDelete.Click += new System.EventHandler(this.cmdDelete_Click);
            // 
            // cmdUpdate
            // 
            this.cmdUpdate.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.cmdUpdate.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdUpdate.Location = new System.Drawing.Point(484, 13);
            this.cmdUpdate.Margin = new System.Windows.Forms.Padding(2);
            this.cmdUpdate.Name = "cmdUpdate";
            this.cmdUpdate.Size = new System.Drawing.Size(68, 30);
            this.cmdUpdate.TabIndex = 3;
            this.cmdUpdate.Text = "Update";
            this.cmdUpdate.UseVisualStyleBackColor = false;
            this.cmdUpdate.Click += new System.EventHandler(this.cmdUpdate_Click);
            // 
            // cmdQuery
            // 
            this.cmdQuery.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.cmdQuery.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdQuery.Location = new System.Drawing.Point(336, 13);
            this.cmdQuery.Margin = new System.Windows.Forms.Padding(2);
            this.cmdQuery.Name = "cmdQuery";
            this.cmdQuery.Size = new System.Drawing.Size(70, 30);
            this.cmdQuery.TabIndex = 2;
            this.cmdQuery.Text = "Query";
            this.cmdQuery.UseVisualStyleBackColor = false;
            this.cmdQuery.Click += new System.EventHandler(this.cmdQuery_Click);
            // 
            // txtRefID
            // 
            this.txtRefID.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtRefID.Location = new System.Drawing.Point(118, 15);
            this.txtRefID.Margin = new System.Windows.Forms.Padding(2);
            this.txtRefID.Name = "txtRefID";
            this.txtRefID.Size = new System.Drawing.Size(198, 23);
            this.txtRefID.TabIndex = 1;
            this.txtRefID.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtRefID_KeyPress);
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(9, 15);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(105, 23);
            this.label1.TabIndex = 0;
            this.label1.Text = "ReferenceID:";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // dgInfo
            // 
            this.dgInfo.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.dgInfo.BackgroundColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgInfo.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgInfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgInfo.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgInfo.Location = new System.Drawing.Point(14, 63);
            this.dgInfo.Margin = new System.Windows.Forms.Padding(2);
            this.dgInfo.Name = "dgInfo";
            this.dgInfo.RowTemplate.Height = 27;
            this.dgInfo.Size = new System.Drawing.Size(646, 368);
            this.dgInfo.TabIndex = 1;
            // 
            // frmUrgentDIDToWH
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(697, 448);
            this.Controls.Add(this.dgInfo);
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "frmUrgentDIDToWH";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmUrgentDIDToWH";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmUrgentDIDToWH_FormClosed);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgInfo)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button cmdQuery;
        private System.Windows.Forms.TextBox txtRefID;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView dgInfo;
        private System.Windows.Forms.Button cmdDelete;
        private System.Windows.Forms.Button cmdUpdate;
    }
}