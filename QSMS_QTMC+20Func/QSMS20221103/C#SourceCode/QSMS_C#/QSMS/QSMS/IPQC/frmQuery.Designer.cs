namespace QSMS.QSMS.IPQC
{
    partial class frmQuery
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
            this.dtGridView = new System.Windows.Forms.DataGridView();
            this.dtEDate = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.dtSDate = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.textDID = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.txtETime = new System.Windows.Forms.TextBox();
            this.txtSTime = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.dtGridView)).BeginInit();
            this.SuspendLayout();
            // 
            // dtGridView
            // 
            this.dtGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dtGridView.Location = new System.Drawing.Point(28, 87);
            this.dtGridView.Name = "dtGridView";
            this.dtGridView.RowTemplate.Height = 23;
            this.dtGridView.Size = new System.Drawing.Size(714, 243);
            this.dtGridView.TabIndex = 33;
            // 
            // dtEDate
            // 
            this.dtEDate.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dtEDate.Location = new System.Drawing.Point(485, 60);
            this.dtEDate.Name = "dtEDate";
            this.dtEDate.Size = new System.Drawing.Size(135, 21);
            this.dtEDate.TabIndex = 32;
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.SystemColors.Info;
            this.label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(388, 60);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(97, 21);
            this.label3.TabIndex = 31;
            this.label3.Text = "结束时间:";
            // 
            // dtSDate
            // 
            this.dtSDate.Font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dtSDate.Location = new System.Drawing.Point(125, 61);
            this.dtSDate.Name = "dtSDate";
            this.dtSDate.Size = new System.Drawing.Size(135, 21);
            this.dtSDate.TabIndex = 30;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.Info;
            this.label2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(28, 61);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(97, 21);
            this.label2.TabIndex = 29;
            this.label2.Text = "开始时间:";
            // 
            // textDID
            // 
            this.textDID.Location = new System.Drawing.Point(131, 25);
            this.textDID.Name = "textDID";
            this.textDID.Size = new System.Drawing.Size(336, 21);
            this.textDID.TabIndex = 28;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(28, 25);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(97, 21);
            this.label1.TabIndex = 27;
            this.label1.Text = "DID或料号:";
            // 
            // button2
            // 
            this.button2.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.button2.Location = new System.Drawing.Point(656, 23);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(86, 26);
            this.button2.TabIndex = 37;
            this.button2.Text = "Execl";
            this.button2.UseVisualStyleBackColor = false;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.button1.Location = new System.Drawing.Point(564, 23);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(86, 26);
            this.button1.TabIndex = 36;
            this.button1.Text = "查询";
            this.button1.UseVisualStyleBackColor = false;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtETime
            // 
            this.txtETime.Location = new System.Drawing.Point(616, 60);
            this.txtETime.Name = "txtETime";
            this.txtETime.Size = new System.Drawing.Size(126, 21);
            this.txtETime.TabIndex = 35;
            // 
            // txtSTime
            // 
            this.txtSTime.Location = new System.Drawing.Point(256, 61);
            this.txtSTime.Name = "txtSTime";
            this.txtSTime.Size = new System.Drawing.Size(126, 21);
            this.txtSTime.TabIndex = 34;
            // 
            // frmQuery
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(768, 363);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.txtETime);
            this.Controls.Add(this.txtSTime);
            this.Controls.Add(this.dtGridView);
            this.Controls.Add(this.dtEDate);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.dtSDate);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textDID);
            this.Controls.Add(this.label1);
            this.Name = "frmQuery";
            this.Text = "frmQuery";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmQuery_FormClosed);
            this.Load += new System.EventHandler(this.frmQuery_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dtGridView)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.DataGridView dtGridView;
        private System.Windows.Forms.DateTimePicker dtEDate;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker dtSDate;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textDID;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txtETime;
        private System.Windows.Forms.TextBox txtSTime;
    }
}