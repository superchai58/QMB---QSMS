namespace QSMS.QSMS.PD
{
    partial class FrmPNCompare
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmPNCompare));
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.textBox4 = new System.Windows.Forms.TextBox();
            this.txtCompPN = new System.Windows.Forms.TextBox();
            this.txtDID = new System.Windows.Forms.TextBox();
            this.textBox10 = new System.Windows.Forms.TextBox();
            this.btn_Excel = new System.Windows.Forms.Button();
            this.btn_Find = new System.Windows.Forms.Button();
            this.GV_Data = new System.Windows.Forms.DataGridView();
            this.txtStatus = new System.Windows.Forms.TextBox();
            this.timer_DID = new System.Windows.Forms.Timer(this.components);
            this.timer_CompPN = new System.Windows.Forms.Timer(this.components);
            this.Begin_Date = new System.Windows.Forms.DateTimePicker();
            this.End_Date = new System.Windows.Forms.DateTimePicker();
            this.dlLine = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.GV_Data)).BeginInit();
            this.SuspendLayout();
            // 
            // textBox1
            // 
            this.textBox1.BackColor = System.Drawing.SystemColors.Info;
            this.textBox1.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox1.Location = new System.Drawing.Point(250, 21);
            this.textBox1.Name = "textBox1";
            this.textBox1.ReadOnly = true;
            this.textBox1.Size = new System.Drawing.Size(87, 19);
            this.textBox1.TabIndex = 0;
            this.textBox1.Text = "Begin Date";
            // 
            // textBox2
            // 
            this.textBox2.BackColor = System.Drawing.SystemColors.Info;
            this.textBox2.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox2.Location = new System.Drawing.Point(492, 18);
            this.textBox2.Name = "textBox2";
            this.textBox2.ReadOnly = true;
            this.textBox2.Size = new System.Drawing.Size(100, 19);
            this.textBox2.TabIndex = 1;
            this.textBox2.Text = "End Date";
            // 
            // textBox3
            // 
            this.textBox3.BackColor = System.Drawing.SystemColors.Info;
            this.textBox3.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox3.Location = new System.Drawing.Point(32, 63);
            this.textBox3.Name = "textBox3";
            this.textBox3.ReadOnly = true;
            this.textBox3.Size = new System.Drawing.Size(69, 19);
            this.textBox3.TabIndex = 2;
            this.textBox3.Text = "DID ";
            // 
            // textBox4
            // 
            this.textBox4.BackColor = System.Drawing.SystemColors.Info;
            this.textBox4.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox4.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox4.Location = new System.Drawing.Point(492, 60);
            this.textBox4.Name = "textBox4";
            this.textBox4.ReadOnly = true;
            this.textBox4.Size = new System.Drawing.Size(69, 19);
            this.textBox4.TabIndex = 3;
            this.textBox4.Text = "CompPN";
            // 
            // txtCompPN
            // 
            this.txtCompPN.Font = new System.Drawing.Font("宋体", 14.25F);
            this.txtCompPN.Location = new System.Drawing.Point(565, 55);
            this.txtCompPN.Name = "txtCompPN";
            this.txtCompPN.Size = new System.Drawing.Size(170, 29);
            this.txtCompPN.TabIndex = 4;
            this.txtCompPN.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtCompPN_KeyDown);
            this.txtCompPN.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCompPN_KeyPress);
            this.txtCompPN.MouseDown += new System.Windows.Forms.MouseEventHandler(this.txtCompPN_MouseDown);
            // 
            // txtDID
            // 
            this.txtDID.Font = new System.Drawing.Font("宋体", 14.25F);
            this.txtDID.Location = new System.Drawing.Point(103, 59);
            this.txtDID.Name = "txtDID";
            this.txtDID.Size = new System.Drawing.Size(377, 29);
            this.txtDID.TabIndex = 7;
            this.txtDID.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtDID_KeyDown);
            this.txtDID.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDID_KeyPress);
            this.txtDID.MouseDown += new System.Windows.Forms.MouseEventHandler(this.txtDID_MouseDown);
            // 
            // textBox10
            // 
            this.textBox10.BackColor = System.Drawing.SystemColors.Info;
            this.textBox10.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.textBox10.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox10.Location = new System.Drawing.Point(32, 24);
            this.textBox10.Name = "textBox10";
            this.textBox10.ReadOnly = true;
            this.textBox10.Size = new System.Drawing.Size(69, 19);
            this.textBox10.TabIndex = 9;
            this.textBox10.Text = "Line";
            // 
            // btn_Excel
            // 
            this.btn_Excel.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btn_Excel.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_Excel.Location = new System.Drawing.Point(753, 55);
            this.btn_Excel.Name = "btn_Excel";
            this.btn_Excel.Size = new System.Drawing.Size(94, 34);
            this.btn_Excel.TabIndex = 10;
            this.btn_Excel.Text = "Excel";
            this.btn_Excel.UseVisualStyleBackColor = false;
            this.btn_Excel.Click += new System.EventHandler(this.btn_Excel_Click);
            // 
            // btn_Find
            // 
            this.btn_Find.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btn_Find.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btn_Find.ForeColor = System.Drawing.SystemColors.ControlText;
            this.btn_Find.Location = new System.Drawing.Point(753, 13);
            this.btn_Find.Name = "btn_Find";
            this.btn_Find.Size = new System.Drawing.Size(94, 36);
            this.btn_Find.TabIndex = 11;
            this.btn_Find.Text = "Find";
            this.btn_Find.UseVisualStyleBackColor = false;
            this.btn_Find.Click += new System.EventHandler(this.btn_Find_Click);
            // 
            // GV_Data
            // 
            this.GV_Data.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.GV_Data.Location = new System.Drawing.Point(32, 108);
            this.GV_Data.Name = "GV_Data";
            this.GV_Data.RowTemplate.Height = 23;
            this.GV_Data.Size = new System.Drawing.Size(815, 316);
            this.GV_Data.TabIndex = 12;
            // 
            // txtStatus
            // 
            this.txtStatus.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.txtStatus.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtStatus.Location = new System.Drawing.Point(32, 430);
            this.txtStatus.Multiline = true;
            this.txtStatus.Name = "txtStatus";
            this.txtStatus.ReadOnly = true;
            this.txtStatus.Size = new System.Drawing.Size(815, 53);
            this.txtStatus.TabIndex = 13;
            // 
            // timer_DID
            // 
            this.timer_DID.Interval = 300;
            this.timer_DID.Tick += new System.EventHandler(this.timer_DID_Tick);
            // 
            // timer_CompPN
            // 
            this.timer_CompPN.Interval = 300;
            this.timer_CompPN.Tick += new System.EventHandler(this.timer_CompPN_Tick);
            // 
            // Begin_Date
            // 
            this.Begin_Date.Font = new System.Drawing.Font("宋体", 14.25F);
            this.Begin_Date.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.Begin_Date.Location = new System.Drawing.Point(339, 16);
            this.Begin_Date.Name = "Begin_Date";
            this.Begin_Date.Size = new System.Drawing.Size(141, 29);
            this.Begin_Date.TabIndex = 14;
            // 
            // End_Date
            // 
            this.End_Date.Font = new System.Drawing.Font("宋体", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.End_Date.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.End_Date.Location = new System.Drawing.Point(594, 14);
            this.End_Date.Name = "End_Date";
            this.End_Date.Size = new System.Drawing.Size(141, 29);
            this.End_Date.TabIndex = 15;
            // 
            // dlLine
            // 
            this.dlLine.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.dlLine.FormattingEnabled = true;
            this.dlLine.Location = new System.Drawing.Point(103, 21);
            this.dlLine.Name = "dlLine";
            this.dlLine.Size = new System.Drawing.Size(141, 24);
            this.dlLine.TabIndex = 16;
            // 
            // FrmPNCompare
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(884, 507);
            this.Controls.Add(this.dlLine);
            this.Controls.Add(this.End_Date);
            this.Controls.Add(this.Begin_Date);
            this.Controls.Add(this.txtStatus);
            this.Controls.Add(this.GV_Data);
            this.Controls.Add(this.btn_Find);
            this.Controls.Add(this.btn_Excel);
            this.Controls.Add(this.textBox10);
            this.Controls.Add(this.txtDID);
            this.Controls.Add(this.txtCompPN);
            this.Controls.Add(this.textBox4);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FrmPNCompare";
            this.Text = "PNCompare";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FrmPNCompare_FormClosed);
            ((System.ComponentModel.ISupportInitialize)(this.GV_Data)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.TextBox textBox4;
        private System.Windows.Forms.TextBox txtCompPN;
        private System.Windows.Forms.TextBox txtDID;
        private System.Windows.Forms.TextBox textBox10;
        private System.Windows.Forms.Button btn_Excel;
        private System.Windows.Forms.Button btn_Find;
        private System.Windows.Forms.DataGridView GV_Data;
        private System.Windows.Forms.TextBox txtStatus;
        private System.Windows.Forms.Timer timer_DID;
        private System.Windows.Forms.Timer timer_CompPN;
        private System.Windows.Forms.DateTimePicker Begin_Date;
        private System.Windows.Forms.DateTimePicker End_Date;
        private System.Windows.Forms.ComboBox dlLine;
    }
}