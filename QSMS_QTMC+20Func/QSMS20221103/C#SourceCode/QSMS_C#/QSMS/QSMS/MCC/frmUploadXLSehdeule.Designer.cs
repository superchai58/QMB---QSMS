namespace QSMS.QSMS.MCC
{
    partial class frmUploadXLSchedule
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
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.btnUpload = new System.Windows.Forms.Button();
            this.btnSelectFile = new System.Windows.Forms.Button();
            this.DataGV = new System.Windows.Forms.DataGridView();
            this.cmsOutput = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.Export = new System.Windows.Forms.ToolStripMenuItem();
            this.gbInfo = new System.Windows.Forms.GroupBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtModel = new System.Windows.Forms.TextBox();
            this.txtNewGroup = new System.Windows.Forms.TextBox();
            this.btnAnalyze = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.txtGroup = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtLine = new System.Windows.Forms.TextBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnGetNewGroup = new System.Windows.Forms.Button();
            this.gvAnalyze = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.txtWO = new System.Windows.Forms.TextBox();
            this.btnSaveSchedule = new System.Windows.Forms.Button();
            this.btnReAssignGroup = new System.Windows.Forms.Button();
            this.gbFunc = new System.Windows.Forms.GroupBox();
            this.txtDate = new System.Windows.Forms.TextBox();
            this.btnOutputSchedule = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.btnClear = new System.Windows.Forms.Button();
            this.btnDelete = new System.Windows.Forms.Button();
            this.btnQueryTmp = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.DataGV)).BeginInit();
            this.cmsOutput.SuspendLayout();
            this.gbInfo.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gvAnalyze)).BeginInit();
            this.gbFunc.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtFilePath
            // 
            this.txtFilePath.Location = new System.Drawing.Point(156, 38);
            this.txtFilePath.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.ReadOnly = true;
            this.txtFilePath.Size = new System.Drawing.Size(846, 29);
            this.txtFilePath.TabIndex = 0;
            // 
            // btnUpload
            // 
            this.btnUpload.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.btnUpload.Location = new System.Drawing.Point(1008, 33);
            this.btnUpload.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnUpload.Name = "btnUpload";
            this.btnUpload.Size = new System.Drawing.Size(112, 34);
            this.btnUpload.TabIndex = 1;
            this.btnUpload.Text = "Upload";
            this.btnUpload.UseVisualStyleBackColor = true;
            this.btnUpload.Click += new System.EventHandler(this.btnUpload_Click);
            // 
            // btnSelectFile
            // 
            this.btnSelectFile.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.btnSelectFile.Location = new System.Drawing.Point(18, 33);
            this.btnSelectFile.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnSelectFile.Name = "btnSelectFile";
            this.btnSelectFile.Size = new System.Drawing.Size(130, 34);
            this.btnSelectFile.TabIndex = 3;
            this.btnSelectFile.Text = "Select File";
            this.btnSelectFile.UseVisualStyleBackColor = true;
            this.btnSelectFile.Click += new System.EventHandler(this.btnSelectFile_Click);
            // 
            // DataGV
            // 
            this.DataGV.AllowUserToAddRows = false;
            this.DataGV.AllowUserToDeleteRows = false;
            this.DataGV.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.DataGV.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.DataGV.ContextMenuStrip = this.cmsOutput;
            this.DataGV.Location = new System.Drawing.Point(20, 78);
            this.DataGV.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.DataGV.Name = "DataGV";
            this.DataGV.ReadOnly = true;
            this.DataGV.RowHeadersVisible = false;
            this.DataGV.RowHeadersWidth = 102;
            this.DataGV.RowTemplate.Height = 24;
            this.DataGV.Size = new System.Drawing.Size(1101, 916);
            this.DataGV.TabIndex = 4;
            this.DataGV.CellMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.DataGV_CellMouseDoubleClick);
            this.DataGV.DataBindingComplete += new System.Windows.Forms.DataGridViewBindingCompleteEventHandler(this.DataGV_DataBindingComplete);
            // 
            // cmsOutput
            // 
            this.cmsOutput.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.cmsOutput.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.Export});
            this.cmsOutput.Name = "cmsOutput";
            this.cmsOutput.Size = new System.Drawing.Size(184, 34);
            this.cmsOutput.Opening += new System.ComponentModel.CancelEventHandler(this.cmsOutput_Opening);
            // 
            // Export
            // 
            this.Export.Name = "Export";
            this.Export.Size = new System.Drawing.Size(183, 30);
            this.Export.Text = "Export Excel";
            this.Export.Click += new System.EventHandler(this.Export_Click);
            // 
            // gbInfo
            // 
            this.gbInfo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gbInfo.Controls.Add(this.label6);
            this.gbInfo.Controls.Add(this.txtModel);
            this.gbInfo.Controls.Add(this.txtNewGroup);
            this.gbInfo.Controls.Add(this.btnAnalyze);
            this.gbInfo.Controls.Add(this.label4);
            this.gbInfo.Controls.Add(this.txtGroup);
            this.gbInfo.Controls.Add(this.label3);
            this.gbInfo.Controls.Add(this.txtLine);
            this.gbInfo.Controls.Add(this.btnSave);
            this.gbInfo.Controls.Add(this.btnGetNewGroup);
            this.gbInfo.Controls.Add(this.gvAnalyze);
            this.gbInfo.Controls.Add(this.label2);
            this.gbInfo.Controls.Add(this.label1);
            this.gbInfo.Controls.Add(this.txtWO);
            this.gbInfo.Location = new System.Drawing.Point(1152, 316);
            this.gbInfo.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.gbInfo.Name = "gbInfo";
            this.gbInfo.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.gbInfo.Size = new System.Drawing.Size(720, 678);
            this.gbInfo.TabIndex = 6;
            this.gbInfo.TabStop = false;
            this.gbInfo.Text = "WO Information";
            // 
            // label6
            // 
            this.label6.AccessibleRole = System.Windows.Forms.AccessibleRole.None;
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label6.Location = new System.Drawing.Point(112, 158);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(80, 24);
            this.label6.TabIndex = 15;
            this.label6.Text = "Model:";
            // 
            // txtModel
            // 
            this.txtModel.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Bold);
            this.txtModel.Location = new System.Drawing.Point(200, 154);
            this.txtModel.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtModel.Name = "txtModel";
            this.txtModel.ReadOnly = true;
            this.txtModel.Size = new System.Drawing.Size(368, 36);
            this.txtModel.TabIndex = 14;
            // 
            // txtNewGroup
            // 
            this.txtNewGroup.BackColor = System.Drawing.Color.LemonChiffon;
            this.txtNewGroup.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Bold);
            this.txtNewGroup.Location = new System.Drawing.Point(202, 300);
            this.txtNewGroup.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtNewGroup.Name = "txtNewGroup";
            this.txtNewGroup.Size = new System.Drawing.Size(362, 36);
            this.txtNewGroup.TabIndex = 13;
            // 
            // btnAnalyze
            // 
            this.btnAnalyze.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Bold);
            this.btnAnalyze.Location = new System.Drawing.Point(444, 248);
            this.btnAnalyze.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnAnalyze.Name = "btnAnalyze";
            this.btnAnalyze.Size = new System.Drawing.Size(126, 40);
            this.btnAnalyze.TabIndex = 12;
            this.btnAnalyze.Text = "Analyze";
            this.btnAnalyze.UseVisualStyleBackColor = true;
            this.btnAnalyze.Click += new System.EventHandler(this.btnAnalyze_Click);
            // 
            // label4
            // 
            this.label4.AccessibleRole = System.Windows.Forms.AccessibleRole.None;
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label4.Location = new System.Drawing.Point(110, 210);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(79, 24);
            this.label4.TabIndex = 11;
            this.label4.Text = "Group:";
            // 
            // txtGroup
            // 
            this.txtGroup.BackColor = System.Drawing.Color.Ivory;
            this.txtGroup.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Bold);
            this.txtGroup.Location = new System.Drawing.Point(200, 204);
            this.txtGroup.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtGroup.Name = "txtGroup";
            this.txtGroup.ReadOnly = true;
            this.txtGroup.Size = new System.Drawing.Size(368, 36);
            this.txtGroup.TabIndex = 10;
            // 
            // label3
            // 
            this.label3.AccessibleRole = System.Windows.Forms.AccessibleRole.None;
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label3.Location = new System.Drawing.Point(126, 102);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(62, 24);
            this.label3.TabIndex = 9;
            this.label3.Text = "Line:";
            // 
            // txtLine
            // 
            this.txtLine.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Bold);
            this.txtLine.Location = new System.Drawing.Point(200, 98);
            this.txtLine.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtLine.Name = "txtLine";
            this.txtLine.ReadOnly = true;
            this.txtLine.Size = new System.Drawing.Size(368, 36);
            this.txtLine.TabIndex = 8;
            // 
            // btnSave
            // 
            this.btnSave.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Bold);
            this.btnSave.Location = new System.Drawing.Point(458, 351);
            this.btnSave.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(112, 34);
            this.btnSave.TabIndex = 7;
            this.btnSave.Text = "Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnGetNewGroup
            // 
            this.btnGetNewGroup.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Bold);
            this.btnGetNewGroup.Location = new System.Drawing.Point(33, 351);
            this.btnGetNewGroup.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnGetNewGroup.Name = "btnGetNewGroup";
            this.btnGetNewGroup.Size = new System.Drawing.Size(148, 34);
            this.btnGetNewGroup.TabIndex = 6;
            this.btnGetNewGroup.Text = "New Group";
            this.btnGetNewGroup.UseVisualStyleBackColor = true;
            this.btnGetNewGroup.Click += new System.EventHandler(this.btnGetNewGroup_Click);
            // 
            // gvAnalyze
            // 
            this.gvAnalyze.AllowUserToAddRows = false;
            this.gvAnalyze.AllowUserToDeleteRows = false;
            this.gvAnalyze.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gvAnalyze.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gvAnalyze.Location = new System.Drawing.Point(33, 408);
            this.gvAnalyze.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.gvAnalyze.Name = "gvAnalyze";
            this.gvAnalyze.ReadOnly = true;
            this.gvAnalyze.RowHeadersVisible = false;
            this.gvAnalyze.RowHeadersWidth = 102;
            this.gvAnalyze.RowTemplate.Height = 24;
            this.gvAnalyze.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.gvAnalyze.Size = new System.Drawing.Size(658, 252);
            this.gvAnalyze.TabIndex = 5;
            this.gvAnalyze.CellMouseDoubleClick += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.gvAnalyze_CellMouseDoubleClick);
            this.gvAnalyze.DataBindingComplete += new System.Windows.Forms.DataGridViewBindingCompleteEventHandler(this.gvAnalyze_DataBindingComplete);
            // 
            // label2
            // 
            this.label2.AccessibleRole = System.Windows.Forms.AccessibleRole.None;
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label2.Location = new System.Drawing.Point(54, 304);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(131, 24);
            this.label2.TabIndex = 3;
            this.label2.Text = "New Group:";
            // 
            // label1
            // 
            this.label1.AccessibleRole = System.Windows.Forms.AccessibleRole.None;
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label1.Location = new System.Drawing.Point(132, 44);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(56, 24);
            this.label1.TabIndex = 1;
            this.label1.Text = "WO:";
            // 
            // txtWO
            // 
            this.txtWO.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Bold);
            this.txtWO.Location = new System.Drawing.Point(200, 39);
            this.txtWO.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtWO.Name = "txtWO";
            this.txtWO.ReadOnly = true;
            this.txtWO.Size = new System.Drawing.Size(368, 36);
            this.txtWO.TabIndex = 0;
            // 
            // btnSaveSchedule
            // 
            this.btnSaveSchedule.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.btnSaveSchedule.ForeColor = System.Drawing.Color.DarkRed;
            this.btnSaveSchedule.Location = new System.Drawing.Point(500, 164);
            this.btnSaveSchedule.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnSaveSchedule.Name = "btnSaveSchedule";
            this.btnSaveSchedule.Size = new System.Drawing.Size(192, 34);
            this.btnSaveSchedule.TabIndex = 7;
            this.btnSaveSchedule.Text = "Save Schedule";
            this.btnSaveSchedule.UseVisualStyleBackColor = true;
            this.btnSaveSchedule.Click += new System.EventHandler(this.btnSaveSchedule_Click);
            // 
            // btnReAssignGroup
            // 
            this.btnReAssignGroup.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.btnReAssignGroup.Location = new System.Drawing.Point(33, 164);
            this.btnReAssignGroup.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnReAssignGroup.Name = "btnReAssignGroup";
            this.btnReAssignGroup.Size = new System.Drawing.Size(192, 34);
            this.btnReAssignGroup.TabIndex = 8;
            this.btnReAssignGroup.Text = "ReAssignGroup";
            this.btnReAssignGroup.UseVisualStyleBackColor = true;
            this.btnReAssignGroup.Click += new System.EventHandler(this.btnReAssignGroup_Click);
            // 
            // gbFunc
            // 
            this.gbFunc.Controls.Add(this.txtDate);
            this.gbFunc.Controls.Add(this.btnOutputSchedule);
            this.gbFunc.Controls.Add(this.label5);
            this.gbFunc.Controls.Add(this.btnClear);
            this.gbFunc.Controls.Add(this.btnDelete);
            this.gbFunc.Controls.Add(this.btnQueryTmp);
            this.gbFunc.Controls.Add(this.btnSaveSchedule);
            this.gbFunc.Controls.Add(this.btnReAssignGroup);
            this.gbFunc.Location = new System.Drawing.Point(1152, 78);
            this.gbFunc.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.gbFunc.Name = "gbFunc";
            this.gbFunc.Padding = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.gbFunc.Size = new System.Drawing.Size(720, 220);
            this.gbFunc.TabIndex = 9;
            this.gbFunc.TabStop = false;
            this.gbFunc.Text = "Function";
            // 
            // txtDate
            // 
            this.txtDate.Font = new System.Drawing.Font("新細明體", 12F, System.Drawing.FontStyle.Bold);
            this.txtDate.Location = new System.Drawing.Point(386, 38);
            this.txtDate.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtDate.Name = "txtDate";
            this.txtDate.Size = new System.Drawing.Size(176, 36);
            this.txtDate.TabIndex = 15;
            // 
            // btnOutputSchedule
            // 
            this.btnOutputSchedule.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Bold);
            this.btnOutputSchedule.Location = new System.Drawing.Point(579, 36);
            this.btnOutputSchedule.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnOutputSchedule.Name = "btnOutputSchedule";
            this.btnOutputSchedule.Size = new System.Drawing.Size(112, 40);
            this.btnOutputSchedule.TabIndex = 14;
            this.btnOutputSchedule.Text = "Output";
            this.btnOutputSchedule.UseVisualStyleBackColor = true;
            this.btnOutputSchedule.Click += new System.EventHandler(this.btnOutputSchedule_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("微軟正黑體", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.label5.Location = new System.Drawing.Point(274, 45);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(111, 29);
            this.label5.TabIndex = 13;
            this.label5.Text = "XL Date :";
            // 
            // btnClear
            // 
            this.btnClear.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.btnClear.Location = new System.Drawing.Point(33, 80);
            this.btnClear.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(192, 34);
            this.btnClear.TabIndex = 11;
            this.btnClear.Text = "Clear";
            this.btnClear.UseVisualStyleBackColor = true;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.btnDelete.Location = new System.Drawing.Point(33, 122);
            this.btnDelete.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(192, 34);
            this.btnDelete.TabIndex = 10;
            this.btnDelete.Text = "Delete Upload";
            this.btnDelete.UseVisualStyleBackColor = true;
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnQueryTmp
            // 
            this.btnQueryTmp.Font = new System.Drawing.Font("微软雅黑", 9F, System.Drawing.FontStyle.Bold);
            this.btnQueryTmp.Location = new System.Drawing.Point(33, 39);
            this.btnQueryTmp.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnQueryTmp.Name = "btnQueryTmp";
            this.btnQueryTmp.Size = new System.Drawing.Size(192, 34);
            this.btnQueryTmp.TabIndex = 9;
            this.btnQueryTmp.Text = "Query";
            this.btnQueryTmp.UseVisualStyleBackColor = true;
            this.btnQueryTmp.Click += new System.EventHandler(this.btnQueryTmp_Click);
            // 
            // frmUploadXLSchedule
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1881, 1082);
            this.Controls.Add(this.gbFunc);
            this.Controls.Add(this.gbInfo);
            this.Controls.Add(this.DataGV);
            this.Controls.Add(this.btnSelectFile);
            this.Controls.Add(this.btnUpload);
            this.Controls.Add(this.txtFilePath);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "frmUploadXLSchedule";
            this.Text = "UploadXLSehdeule_AutoGroup 20230308";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmUploadXLSchedule_FormClosed);
            this.Load += new System.EventHandler(this.frmUploadXLSchedule_Load);
            ((System.ComponentModel.ISupportInitialize)(this.DataGV)).EndInit();
            this.cmsOutput.ResumeLayout(false);
            this.gbInfo.ResumeLayout(false);
            this.gbInfo.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gvAnalyze)).EndInit();
            this.gbFunc.ResumeLayout(false);
            this.gbFunc.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Button btnUpload;
        private System.Windows.Forms.Button btnSelectFile;
        private System.Windows.Forms.DataGridView DataGV;
        private System.Windows.Forms.GroupBox gbInfo;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtGroup;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtLine;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnGetNewGroup;
        private System.Windows.Forms.DataGridView gvAnalyze;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtWO;
        private System.Windows.Forms.Button btnAnalyze;
        private System.Windows.Forms.TextBox txtNewGroup;
        private System.Windows.Forms.Button btnSaveSchedule;
        private System.Windows.Forms.Button btnReAssignGroup;
        private System.Windows.Forms.GroupBox gbFunc;
        private System.Windows.Forms.Button btnDelete;
        private System.Windows.Forms.Button btnQueryTmp;
        private System.Windows.Forms.Button btnClear;
        private System.Windows.Forms.ContextMenuStrip cmsOutput;
        private System.Windows.Forms.ToolStripMenuItem Export;
        private System.Windows.Forms.TextBox txtDate;
        private System.Windows.Forms.Button btnOutputSchedule;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtModel;
    }
}