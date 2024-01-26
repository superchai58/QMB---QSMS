namespace QSMS.QSMS.SpecialCase
{
    partial class frmUnChkWO
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
            this.txtWO = new System.Windows.Forms.TextBox();
            this.btnAdd = new System.Windows.Forms.Button();
            this.btnDel = new System.Windows.Forms.Button();
            this.dgInfo = new System.Windows.Forms.DataGridView();
            this.label2 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dgInfo)).BeginInit();
            this.SuspendLayout();
            // 
            // txtWO
            // 
            this.txtWO.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.txtWO.Location = new System.Drawing.Point(132, 51);
            this.txtWO.Margin = new System.Windows.Forms.Padding(2);
            this.txtWO.Name = "txtWO";
            this.txtWO.Size = new System.Drawing.Size(278, 26);
            this.txtWO.TabIndex = 27;
            this.txtWO.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtWO_KeyPress);
            // 
            // btnAdd
            // 
            this.btnAdd.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnAdd.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnAdd.Location = new System.Drawing.Point(413, 50);
            this.btnAdd.Margin = new System.Windows.Forms.Padding(2);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(70, 30);
            this.btnAdd.TabIndex = 30;
            this.btnAdd.Text = "Add";
            this.btnAdd.UseVisualStyleBackColor = false;
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // btnDel
            // 
            this.btnDel.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.btnDel.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.btnDel.Location = new System.Drawing.Point(498, 50);
            this.btnDel.Margin = new System.Windows.Forms.Padding(2);
            this.btnDel.Name = "btnDel";
            this.btnDel.Size = new System.Drawing.Size(70, 30);
            this.btnDel.TabIndex = 31;
            this.btnDel.Text = "Delete";
            this.btnDel.UseVisualStyleBackColor = false;
            this.btnDel.Click += new System.EventHandler(this.btnDel_Click);
            // 
            // dgInfo
            // 
            this.dgInfo.BackgroundColor = System.Drawing.Color.White;
            this.dgInfo.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgInfo.GridColor = System.Drawing.Color.White;
            this.dgInfo.Location = new System.Drawing.Point(20, 97);
            this.dgInfo.Margin = new System.Windows.Forms.Padding(2);
            this.dgInfo.Name = "dgInfo";
            this.dgInfo.RowTemplate.Height = 27;
            this.dgInfo.Size = new System.Drawing.Size(597, 328);
            this.dgInfo.TabIndex = 32;
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.Info;
            this.label2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(29, 50);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(98, 26);
            this.label2.TabIndex = 33;
            this.label2.Text = "Work_Order:";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // frmUnChkWO
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(667, 453);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dgInfo);
            this.Controls.Add(this.btnDel);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.txtWO);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "frmUnChkWO";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmUnChkWO";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmUnChkWO_FormClosed);
            this.Load += new System.EventHandler(this.frmUnChkWO_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgInfo)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txtWO;
        private System.Windows.Forms.Button btnAdd;
        private System.Windows.Forms.Button btnDel;
        private System.Windows.Forms.DataGridView dgInfo;
        private System.Windows.Forms.Label label2;
    }
}