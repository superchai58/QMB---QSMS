namespace QSMS.QSMS.MCC
{
    partial class frmTransferFujiXML
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
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.txtFile = new System.Windows.Forms.TextBox();
            this.lstFile2 = new System.Windows.Forms.ListView();
            this.columnHeader3 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.lstFile1 = new System.Windows.Forms.ListView();
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.cmdDELALL = new System.Windows.Forms.Button();
            this.cmdADDALL = new System.Windows.Forms.Button();
            this.cmdDEL = new System.Windows.Forms.Button();
            this.cmdADD = new System.Windows.Forms.Button();
            this.textBox6 = new System.Windows.Forms.TextBox();
            this.textBox5 = new System.Windows.Forms.TextBox();
            this.btnUpload = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.lblMsg = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.BackColor = System.Drawing.SystemColors.Info;
            this.btnSelectFile.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnSelectFile.ForeColor = System.Drawing.SystemColors.WindowText;
            this.btnSelectFile.Location = new System.Drawing.Point(35, 20);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(86, 27);
            this.btnSelectFile.TabIndex = 6;
            this.btnSelectFile.Text = "浏览";
            this.btnSelectFile.UseVisualStyleBackColor = false;
            this.btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click);
            // 
            // txtFile
            // 
            this.txtFile.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtFile.Location = new System.Drawing.Point(126, 21);
            this.txtFile.Margin = new System.Windows.Forms.Padding(2);
            this.txtFile.Name = "txtFile";
            this.txtFile.Size = new System.Drawing.Size(707, 26);
            this.txtFile.TabIndex = 7;
            // 
            // lstFile2
            // 
            this.lstFile2.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader3});
            this.lstFile2.Location = new System.Drawing.Point(487, 125);
            this.lstFile2.Name = "lstFile2";
            this.lstFile2.Size = new System.Drawing.Size(346, 308);
            this.lstFile2.TabIndex = 46;
            this.lstFile2.UseCompatibleStateImageBehavior = false;
            this.lstFile2.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader3
            // 
            this.columnHeader3.Text = "FileName";
            this.columnHeader3.Width = 335;
            // 
            // lstFile1
            // 
            this.lstFile1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader2});
            this.lstFile1.Location = new System.Drawing.Point(25, 125);
            this.lstFile1.Name = "lstFile1";
            this.lstFile1.Size = new System.Drawing.Size(359, 308);
            this.lstFile1.TabIndex = 45;
            this.lstFile1.UseCompatibleStateImageBehavior = false;
            this.lstFile1.View = System.Windows.Forms.View.Details;
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "FileName";
            this.columnHeader2.Width = 335;
            // 
            // cmdDELALL
            // 
            this.cmdDELALL.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdDELALL.ForeColor = System.Drawing.Color.Red;
            this.cmdDELALL.Location = new System.Drawing.Point(410, 305);
            this.cmdDELALL.Name = "cmdDELALL";
            this.cmdDELALL.Size = new System.Drawing.Size(51, 29);
            this.cmdDELALL.TabIndex = 43;
            this.cmdDELALL.Text = "<<";
            this.cmdDELALL.UseVisualStyleBackColor = true;
            this.cmdDELALL.Click += new System.EventHandler(this.cmdDELALL_Click);
            // 
            // cmdADDALL
            // 
            this.cmdADDALL.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdADDALL.ForeColor = System.Drawing.Color.Blue;
            this.cmdADDALL.Location = new System.Drawing.Point(410, 205);
            this.cmdADDALL.Name = "cmdADDALL";
            this.cmdADDALL.Size = new System.Drawing.Size(51, 29);
            this.cmdADDALL.TabIndex = 41;
            this.cmdADDALL.Text = ">>";
            this.cmdADDALL.UseVisualStyleBackColor = true;
            this.cmdADDALL.Click += new System.EventHandler(this.cmdADDALL_Click);
            // 
            // cmdDEL
            // 
            this.cmdDEL.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdDEL.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(128)))), ((int)(((byte)(128)))));
            this.cmdDEL.Location = new System.Drawing.Point(410, 254);
            this.cmdDEL.Name = "cmdDEL";
            this.cmdDEL.Size = new System.Drawing.Size(51, 29);
            this.cmdDEL.TabIndex = 42;
            this.cmdDEL.Text = "<";
            this.cmdDEL.UseVisualStyleBackColor = true;
            this.cmdDEL.Click += new System.EventHandler(this.cmdDEL_Click);
            // 
            // cmdADD
            // 
            this.cmdADD.Font = new System.Drawing.Font("微软雅黑", 10.5F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdADD.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(128)))), ((int)(((byte)(128)))), ((int)(((byte)(255)))));
            this.cmdADD.Location = new System.Drawing.Point(410, 155);
            this.cmdADD.Name = "cmdADD";
            this.cmdADD.Size = new System.Drawing.Size(51, 29);
            this.cmdADD.TabIndex = 40;
            this.cmdADD.Text = ">";
            this.cmdADD.UseVisualStyleBackColor = true;
            this.cmdADD.Click += new System.EventHandler(this.cmdADD_Click);
            // 
            // textBox6
            // 
            this.textBox6.BackColor = System.Drawing.SystemColors.Info;
            this.textBox6.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox6.ForeColor = System.Drawing.Color.Black;
            this.textBox6.Location = new System.Drawing.Point(487, 96);
            this.textBox6.Name = "textBox6";
            this.textBox6.ReadOnly = true;
            this.textBox6.Size = new System.Drawing.Size(346, 23);
            this.textBox6.TabIndex = 44;
            this.textBox6.Text = "NeedUpload";
            this.textBox6.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // textBox5
            // 
            this.textBox5.BackColor = System.Drawing.SystemColors.Info;
            this.textBox5.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.textBox5.ForeColor = System.Drawing.Color.Black;
            this.textBox5.Location = new System.Drawing.Point(25, 96);
            this.textBox5.Name = "textBox5";
            this.textBox5.ReadOnly = true;
            this.textBox5.Size = new System.Drawing.Size(359, 23);
            this.textBox5.TabIndex = 39;
            this.textBox5.Text = "FileList";
            this.textBox5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // btnUpload
            // 
            this.btnUpload.BackColor = System.Drawing.SystemColors.Info;
            this.btnUpload.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnUpload.ForeColor = System.Drawing.SystemColors.InactiveCaptionText;
            this.btnUpload.Location = new System.Drawing.Point(759, 51);
            this.btnUpload.Margin = new System.Windows.Forms.Padding(2);
            this.btnUpload.Name = "btnUpload";
            this.btnUpload.Size = new System.Drawing.Size(74, 40);
            this.btnUpload.TabIndex = 47;
            this.btnUpload.Text = "上传";
            this.btnUpload.UseVisualStyleBackColor = false;
            this.btnUpload.Click += new System.EventHandler(this.btnUpload_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.groupBox2.Controls.Add(this.lblMsg);
            this.groupBox2.Font = new System.Drawing.Font("宋体", 10.5F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.groupBox2.Location = new System.Drawing.Point(25, 438);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox2.Size = new System.Drawing.Size(808, 64);
            this.groupBox2.TabIndex = 48;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Message";
            // 
            // lblMsg
            // 
            this.lblMsg.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(224)))), ((int)(((byte)(192)))));
            this.lblMsg.Font = new System.Drawing.Font("宋体", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblMsg.ForeColor = System.Drawing.Color.Red;
            this.lblMsg.Location = new System.Drawing.Point(4, 16);
            this.lblMsg.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(788, 31);
            this.lblMsg.TabIndex = 4;
            this.lblMsg.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // frmTransferFujiXML
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(854, 513);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.btnUpload);
            this.Controls.Add(this.lstFile2);
            this.Controls.Add(this.lstFile1);
            this.Controls.Add(this.cmdDELALL);
            this.Controls.Add(this.cmdADDALL);
            this.Controls.Add(this.cmdDEL);
            this.Controls.Add(this.cmdADD);
            this.Controls.Add(this.textBox6);
            this.Controls.Add(this.textBox5);
            this.Controls.Add(this.btnSelectFile);
            this.Controls.Add(this.txtFile);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "frmTransferFujiXML";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmTransferFujiXML 20230323";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmTransferFujiXML_FormClosed);
            this.Load += new System.EventHandler(this.frmTransferFujiXML_Load);
            this.groupBox2.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.TextBox txtFile;
        private System.Windows.Forms.ListView lstFile2;
        private System.Windows.Forms.ColumnHeader columnHeader3;
        private System.Windows.Forms.ListView lstFile1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
        private System.Windows.Forms.Button cmdDELALL;
        private System.Windows.Forms.Button cmdADDALL;
        private System.Windows.Forms.Button cmdDEL;
        private System.Windows.Forms.Button cmdADD;
        private System.Windows.Forms.TextBox textBox6;
        private System.Windows.Forms.TextBox textBox5;
        private System.Windows.Forms.Button btnUpload;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label lblMsg;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
    }
}