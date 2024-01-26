namespace QSMS.QSMS.SpecialCase
{
    partial class frmStartSplitLineMC
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
            this.CmdStartSplitLineMC = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // CmdStartSplitLineMC
            // 
            this.CmdStartSplitLineMC.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.CmdStartSplitLineMC.Font = new System.Drawing.Font("宋体", 18F);
            this.CmdStartSplitLineMC.Location = new System.Drawing.Point(127, 78);
            this.CmdStartSplitLineMC.Margin = new System.Windows.Forms.Padding(2);
            this.CmdStartSplitLineMC.Name = "CmdStartSplitLineMC";
            this.CmdStartSplitLineMC.Size = new System.Drawing.Size(164, 46);
            this.CmdStartSplitLineMC.TabIndex = 1;
            this.CmdStartSplitLineMC.Text = "启 动";
            this.CmdStartSplitLineMC.UseVisualStyleBackColor = false;
            this.CmdStartSplitLineMC.Click += new System.EventHandler(this.CmdStartSplitLineMC_Click);
            // 
            // frmStartSplitLineMC
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(424, 210);
            this.Controls.Add(this.CmdStartSplitLineMC);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "frmStartSplitLineMC";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "启动分仓";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.frmStartSplitLineMC_FormClosed);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button CmdStartSplitLineMC;
    }
}