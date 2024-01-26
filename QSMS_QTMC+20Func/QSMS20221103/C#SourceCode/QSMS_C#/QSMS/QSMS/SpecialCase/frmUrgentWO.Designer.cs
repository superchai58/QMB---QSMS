namespace QSMS.QSMS.SpecialCase
{
    partial class frmUrgentWO
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
            this.lblWO = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.txtWO = new System.Windows.Forms.TextBox();
            this.btnQuery = new System.Windows.Forms.Button();
            this.btnUrgent = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dgWOPlanSeq = new System.Windows.Forms.DataGridView();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.dgCurWoSeq = new System.Windows.Forms.DataGridView();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.dgWoInputPlan = new System.Windows.Forms.DataGridView();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.dgQSMS_WO_XL = new System.Windows.Forms.DataGridView();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgWOPlanSeq)).BeginInit();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgCurWoSeq)).BeginInit();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgWoInputPlan)).BeginInit();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgQSMS_WO_XL)).BeginInit();
            this.SuspendLayout();
            // 
            // lblWO
            // 
            this.lblWO.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.lblWO.Location = new System.Drawing.Point(669, 68);
            this.lblWO.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblWO.Name = "lblWO";
            this.lblWO.Size = new System.Drawing.Size(89, 23);
            this.lblWO.TabIndex = 4;
            this.lblWO.Text = "WorkOrder:";
            this.lblWO.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.label2.Location = new System.Drawing.Point(661, 135);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(144, 47);
            this.label2.TabIndex = 5;
            this.label2.Text = "输入工单，点击Query按钮，确定工单的信息正后,点击紧急插单按钮。";
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.label3.ForeColor = System.Drawing.Color.Red;
            this.label3.Location = new System.Drawing.Point(669, 274);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(122, 53);
            this.label3.TabIndex = 6;
            this.label3.Text = "紧急插单只需要输入大板的工单号码，小板会随着大板一起插单。";
            // 
            // txtWO
            // 
            this.txtWO.Location = new System.Drawing.Point(669, 93);
            this.txtWO.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.txtWO.Name = "txtWO";
            this.txtWO.Size = new System.Drawing.Size(137, 21);
            this.txtWO.TabIndex = 7;
            this.txtWO.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtWO_KeyPress);
            // 
            // btnQuery
            // 
            this.btnQuery.AutoSize = true;
            this.btnQuery.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnQuery.Location = new System.Drawing.Point(688, 184);
            this.btnQuery.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnQuery.Name = "btnQuery";
            this.btnQuery.Size = new System.Drawing.Size(85, 30);
            this.btnQuery.TabIndex = 8;
            this.btnQuery.Text = "查询";
            this.btnQuery.UseVisualStyleBackColor = false;
            this.btnQuery.Click += new System.EventHandler(this.btnQuery_Click);
            // 
            // btnUrgent
            // 
            this.btnUrgent.AutoSize = true;
            this.btnUrgent.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnUrgent.Location = new System.Drawing.Point(688, 340);
            this.btnUrgent.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnUrgent.Name = "btnUrgent";
            this.btnUrgent.Size = new System.Drawing.Size(85, 30);
            this.btnUrgent.TabIndex = 9;
            this.btnUrgent.Text = "紧急插单";
            this.btnUrgent.UseVisualStyleBackColor = false;
            this.btnUrgent.Click += new System.EventHandler(this.btnUrgent_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.dgWOPlanSeq);
            this.groupBox1.Location = new System.Drawing.Point(9, 25);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox1.Size = new System.Drawing.Size(608, 116);
            this.groupBox1.TabIndex = 10;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "XL_WOPlanSeq";
            // 
            // dgWOPlanSeq
            // 
            this.dgWOPlanSeq.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.ColumnHeader;
            this.dgWOPlanSeq.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllHeaders;
            this.dgWOPlanSeq.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgWOPlanSeq.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgWOPlanSeq.Location = new System.Drawing.Point(2, 16);
            this.dgWOPlanSeq.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.dgWOPlanSeq.Name = "dgWOPlanSeq";
            this.dgWOPlanSeq.RowTemplate.Height = 27;
            this.dgWOPlanSeq.Size = new System.Drawing.Size(604, 98);
            this.dgWOPlanSeq.TabIndex = 0;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.dgCurWoSeq);
            this.groupBox2.Location = new System.Drawing.Point(9, 146);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox2.Size = new System.Drawing.Size(605, 123);
            this.groupBox2.TabIndex = 11;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "XL_CurWOSeq";
            // 
            // dgCurWoSeq
            // 
            this.dgCurWoSeq.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.ColumnHeader;
            this.dgCurWoSeq.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllHeaders;
            this.dgCurWoSeq.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgCurWoSeq.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgCurWoSeq.Location = new System.Drawing.Point(2, 16);
            this.dgCurWoSeq.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.dgCurWoSeq.Name = "dgCurWoSeq";
            this.dgCurWoSeq.RowTemplate.Height = 27;
            this.dgCurWoSeq.Size = new System.Drawing.Size(601, 105);
            this.dgCurWoSeq.TabIndex = 1;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.dgWoInputPlan);
            this.groupBox3.Location = new System.Drawing.Point(9, 274);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox3.Size = new System.Drawing.Size(603, 124);
            this.groupBox3.TabIndex = 12;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "QSMS_WoInputPlan";
            // 
            // dgWoInputPlan
            // 
            this.dgWoInputPlan.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.ColumnHeader;
            this.dgWoInputPlan.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllHeaders;
            this.dgWoInputPlan.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgWoInputPlan.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgWoInputPlan.Location = new System.Drawing.Point(2, 16);
            this.dgWoInputPlan.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.dgWoInputPlan.Name = "dgWoInputPlan";
            this.dgWoInputPlan.RowTemplate.Height = 27;
            this.dgWoInputPlan.Size = new System.Drawing.Size(599, 106);
            this.dgWoInputPlan.TabIndex = 1;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.dgQSMS_WO_XL);
            this.groupBox4.Location = new System.Drawing.Point(9, 402);
            this.groupBox4.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.groupBox4.Size = new System.Drawing.Size(601, 136);
            this.groupBox4.TabIndex = 13;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "QSMS_WO_XL";
            // 
            // dgQSMS_WO_XL
            // 
            this.dgQSMS_WO_XL.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.ColumnHeader;
            this.dgQSMS_WO_XL.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllHeaders;
            this.dgQSMS_WO_XL.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgQSMS_WO_XL.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dgQSMS_WO_XL.Location = new System.Drawing.Point(2, 16);
            this.dgQSMS_WO_XL.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.dgQSMS_WO_XL.Name = "dgQSMS_WO_XL";
            this.dgQSMS_WO_XL.RowTemplate.Height = 27;
            this.dgQSMS_WO_XL.Size = new System.Drawing.Size(597, 118);
            this.dgQSMS_WO_XL.TabIndex = 1;
            // 
            // frmUrgentWO
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(854, 548);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.btnUrgent);
            this.Controls.Add(this.btnQuery);
            this.Controls.Add(this.txtWO);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lblWO);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "frmUrgentWO";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmUrgentWO";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmUrgentWO_FormClosed);
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgWOPlanSeq)).EndInit();
            this.groupBox2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgCurWoSeq)).EndInit();
            this.groupBox3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgWoInputPlan)).EndInit();
            this.groupBox4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgQSMS_WO_XL)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblWO;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtWO;
        private System.Windows.Forms.Button btnQuery;
        private System.Windows.Forms.Button btnUrgent;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.DataGridView dgWOPlanSeq;
        private System.Windows.Forms.DataGridView dgCurWoSeq;
        private System.Windows.Forms.DataGridView dgWoInputPlan;
        private System.Windows.Forms.DataGridView dgQSMS_WO_XL;
    }
}