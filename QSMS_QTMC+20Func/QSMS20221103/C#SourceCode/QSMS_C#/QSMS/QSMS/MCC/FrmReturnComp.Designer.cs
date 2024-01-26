namespace QSMS.QSMS.MCC
{
    partial class FrmReturnComp
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmReturnComp));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.OptSATO = new System.Windows.Forms.RadioButton();
            this.OptZebra = new System.Windows.Forms.RadioButton();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.optNetwork = new System.Windows.Forms.RadioButton();
            this.OptPrint = new System.Windows.Forms.RadioButton();
            this.OptComp = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtCompPort = new System.Windows.Forms.TextBox();
            this.txtComm = new System.Windows.Forms.TextBox();
            this.CmdCommSave = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.cmdGetRefID = new System.Windows.Forms.Button();
            this.cmdReprint = new System.Windows.Forms.Button();
            this.cmdOK = new System.Windows.Forms.Button();
            this.ChkBGA = new System.Windows.Forms.CheckBox();
            this.ChkHUA = new System.Windows.Forms.CheckBox();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.optBadMaterial = new System.Windows.Forms.RadioButton();
            this.optGoodMaterial = new System.Windows.Forms.RadioButton();
            this.txtLotCode = new System.Windows.Forms.TextBox();
            this.txtVendorCode = new System.Windows.Forms.TextBox();
            this.txtQty = new System.Windows.Forms.TextBox();
            this.txtDateCode = new System.Windows.Forms.TextBox();
            this.txtCompPN = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.gridReturnComp = new System.Windows.Forms.DataGridView();
            this.LblMessage = new System.Windows.Forms.Label();
            this.lblFeedBack = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.gridReturnComp)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox1.Controls.Add(this.OptSATO);
            this.groupBox1.Controls.Add(this.OptZebra);
            this.groupBox1.Location = new System.Drawing.Point(7, 23);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox1.Size = new System.Drawing.Size(226, 33);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            // 
            // OptSATO
            // 
            this.OptSATO.AutoSize = true;
            this.OptSATO.BackColor = System.Drawing.SystemColors.Info;
            this.OptSATO.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold);
            this.OptSATO.ForeColor = System.Drawing.Color.DarkRed;
            this.OptSATO.Location = new System.Drawing.Point(115, 10);
            this.OptSATO.Margin = new System.Windows.Forms.Padding(2);
            this.OptSATO.Name = "OptSATO";
            this.OptSATO.Size = new System.Drawing.Size(107, 19);
            this.OptSATO.TabIndex = 1;
            this.OptSATO.Text = "SATO Printer";
            this.OptSATO.UseVisualStyleBackColor = false;
            // 
            // OptZebra
            // 
            this.OptZebra.AutoSize = true;
            this.OptZebra.BackColor = System.Drawing.SystemColors.Info;
            this.OptZebra.Checked = true;
            this.OptZebra.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.OptZebra.ForeColor = System.Drawing.Color.Green;
            this.OptZebra.Location = new System.Drawing.Point(9, 10);
            this.OptZebra.Margin = new System.Windows.Forms.Padding(2);
            this.OptZebra.Name = "OptZebra";
            this.OptZebra.Size = new System.Drawing.Size(109, 19);
            this.OptZebra.TabIndex = 0;
            this.OptZebra.TabStop = true;
            this.OptZebra.Text = "Zebra Printer";
            this.OptZebra.UseVisualStyleBackColor = false;
            // 
            // groupBox2
            // 
            this.groupBox2.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox2.Controls.Add(this.optNetwork);
            this.groupBox2.Controls.Add(this.OptPrint);
            this.groupBox2.Controls.Add(this.OptComp);
            this.groupBox2.Location = new System.Drawing.Point(248, 23);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox2.Size = new System.Drawing.Size(246, 33);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            // 
            // optNetwork
            // 
            this.optNetwork.AutoSize = true;
            this.optNetwork.BackColor = System.Drawing.SystemColors.Info;
            this.optNetwork.Checked = true;
            this.optNetwork.Location = new System.Drawing.Point(170, 10);
            this.optNetwork.Margin = new System.Windows.Forms.Padding(2);
            this.optNetwork.Name = "optNetwork";
            this.optNetwork.Size = new System.Drawing.Size(65, 16);
            this.optNetwork.TabIndex = 2;
            this.optNetwork.TabStop = true;
            this.optNetwork.Text = "Network";
            this.optNetwork.UseVisualStyleBackColor = false;
            // 
            // OptPrint
            // 
            this.OptPrint.AutoSize = true;
            this.OptPrint.BackColor = System.Drawing.SystemColors.Info;
            this.OptPrint.Location = new System.Drawing.Point(86, 10);
            this.OptPrint.Margin = new System.Windows.Forms.Padding(2);
            this.OptPrint.Name = "OptPrint";
            this.OptPrint.Size = new System.Drawing.Size(83, 16);
            this.OptPrint.TabIndex = 1;
            this.OptPrint.Text = "Print Port";
            this.OptPrint.UseVisualStyleBackColor = false;
            // 
            // OptComp
            // 
            this.OptComp.AutoSize = true;
            this.OptComp.BackColor = System.Drawing.SystemColors.Info;
            this.OptComp.Location = new System.Drawing.Point(9, 10);
            this.OptComp.Margin = new System.Windows.Forms.Padding(2);
            this.OptComp.Name = "OptComp";
            this.OptComp.Size = new System.Drawing.Size(77, 16);
            this.OptComp.TabIndex = 0;
            this.OptComp.Text = "Comp Port";
            this.OptComp.UseVisualStyleBackColor = false;
            // 
            // label1
            // 
            this.label1.BackColor = System.Drawing.SystemColors.Info;
            this.label1.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label1.Location = new System.Drawing.Point(523, 16);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(90, 24);
            this.label1.TabIndex = 2;
            this.label1.Text = "Comp Port";
            // 
            // label2
            // 
            this.label2.BackColor = System.Drawing.SystemColors.Info;
            this.label2.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label2.Location = new System.Drawing.Point(526, 47);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(87, 23);
            this.label2.TabIndex = 3;
            this.label2.Text = "Settings";
            // 
            // txtCompPort
            // 
            this.txtCompPort.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.5F, System.Drawing.FontStyle.Bold);
            this.txtCompPort.Location = new System.Drawing.Point(617, 17);
            this.txtCompPort.Margin = new System.Windows.Forms.Padding(2);
            this.txtCompPort.Name = "txtCompPort";
            this.txtCompPort.Size = new System.Drawing.Size(118, 23);
            this.txtCompPort.TabIndex = 4;
            this.txtCompPort.Text = "1";
            // 
            // txtComm
            // 
            this.txtComm.Font = new System.Drawing.Font("Microsoft Sans Serif", 10.5F, System.Drawing.FontStyle.Bold);
            this.txtComm.Location = new System.Drawing.Point(617, 47);
            this.txtComm.Margin = new System.Windows.Forms.Padding(2);
            this.txtComm.Name = "txtComm";
            this.txtComm.Size = new System.Drawing.Size(118, 23);
            this.txtComm.TabIndex = 5;
            this.txtComm.Text = "9600,N,8,1";
            // 
            // CmdCommSave
            // 
            this.CmdCommSave.Image = ((System.Drawing.Image)(resources.GetObject("CmdCommSave.Image")));
            this.CmdCommSave.Location = new System.Drawing.Point(758, 16);
            this.CmdCommSave.Margin = new System.Windows.Forms.Padding(2);
            this.CmdCommSave.Name = "CmdCommSave";
            this.CmdCommSave.Size = new System.Drawing.Size(82, 58);
            this.CmdCommSave.TabIndex = 6;
            this.CmdCommSave.Text = "Comm Save";
            this.CmdCommSave.TextAlign = System.Drawing.ContentAlignment.BottomCenter;
            this.CmdCommSave.UseVisualStyleBackColor = true;
            this.CmdCommSave.Click += new System.EventHandler(this.CmdCommSave_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.groupBox3.Controls.Add(this.CmdCommSave);
            this.groupBox3.Controls.Add(this.txtComm);
            this.groupBox3.Controls.Add(this.txtCompPort);
            this.groupBox3.Controls.Add(this.label2);
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Controls.Add(this.groupBox2);
            this.groupBox3.Controls.Add(this.groupBox1);
            this.groupBox3.Location = new System.Drawing.Point(10, 9);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox3.Size = new System.Drawing.Size(904, 83);
            this.groupBox3.TabIndex = 7;
            this.groupBox3.TabStop = false;
            // 
            // groupBox4
            // 
            this.groupBox4.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.groupBox4.Controls.Add(this.cmdGetRefID);
            this.groupBox4.Controls.Add(this.cmdReprint);
            this.groupBox4.Controls.Add(this.cmdOK);
            this.groupBox4.Controls.Add(this.ChkBGA);
            this.groupBox4.Controls.Add(this.ChkHUA);
            this.groupBox4.Controls.Add(this.groupBox5);
            this.groupBox4.Controls.Add(this.txtLotCode);
            this.groupBox4.Controls.Add(this.txtVendorCode);
            this.groupBox4.Controls.Add(this.txtQty);
            this.groupBox4.Controls.Add(this.txtDateCode);
            this.groupBox4.Controls.Add(this.txtCompPN);
            this.groupBox4.Controls.Add(this.label7);
            this.groupBox4.Controls.Add(this.label6);
            this.groupBox4.Controls.Add(this.label5);
            this.groupBox4.Controls.Add(this.label4);
            this.groupBox4.Controls.Add(this.label3);
            this.groupBox4.Location = new System.Drawing.Point(10, 96);
            this.groupBox4.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox4.Size = new System.Drawing.Size(904, 115);
            this.groupBox4.TabIndex = 8;
            this.groupBox4.TabStop = false;
            // 
            // cmdGetRefID
            // 
            this.cmdGetRefID.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdGetRefID.Location = new System.Drawing.Point(489, 76);
            this.cmdGetRefID.Margin = new System.Windows.Forms.Padding(2);
            this.cmdGetRefID.Name = "cmdGetRefID";
            this.cmdGetRefID.Size = new System.Drawing.Size(109, 30);
            this.cmdGetRefID.TabIndex = 17;
            this.cmdGetRefID.Text = "GetRefID";
            this.cmdGetRefID.UseVisualStyleBackColor = true;
            this.cmdGetRefID.Click += new System.EventHandler(this.cmdGetRefID_Click);
            // 
            // cmdReprint
            // 
            this.cmdReprint.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdReprint.Location = new System.Drawing.Point(381, 76);
            this.cmdReprint.Margin = new System.Windows.Forms.Padding(2);
            this.cmdReprint.Name = "cmdReprint";
            this.cmdReprint.Size = new System.Drawing.Size(104, 30);
            this.cmdReprint.TabIndex = 16;
            this.cmdReprint.Text = "Reprint";
            this.cmdReprint.UseVisualStyleBackColor = true;
            this.cmdReprint.Click += new System.EventHandler(this.cmdReprint_Click);
            // 
            // cmdOK
            // 
            this.cmdOK.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.cmdOK.Location = new System.Drawing.Point(301, 76);
            this.cmdOK.Margin = new System.Windows.Forms.Padding(2);
            this.cmdOK.Name = "cmdOK";
            this.cmdOK.Size = new System.Drawing.Size(76, 30);
            this.cmdOK.TabIndex = 15;
            this.cmdOK.Text = "OK";
            this.cmdOK.UseVisualStyleBackColor = true;
            this.cmdOK.Click += new System.EventHandler(this.cmdOK_Click);
            // 
            // ChkBGA
            // 
            this.ChkBGA.BackColor = System.Drawing.SystemColors.Info;
            this.ChkBGA.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ChkBGA.Location = new System.Drawing.Point(808, 20);
            this.ChkBGA.Margin = new System.Windows.Forms.Padding(2);
            this.ChkBGA.Name = "ChkBGA";
            this.ChkBGA.Size = new System.Drawing.Size(72, 19);
            this.ChkBGA.TabIndex = 14;
            this.ChkBGA.Text = "重置球";
            this.ChkBGA.UseVisualStyleBackColor = false;
            // 
            // ChkHUA
            // 
            this.ChkHUA.BackColor = System.Drawing.SystemColors.Info;
            this.ChkHUA.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.ChkHUA.Location = new System.Drawing.Point(737, 18);
            this.ChkHUA.Margin = new System.Windows.Forms.Padding(2);
            this.ChkHUA.Name = "ChkHUA";
            this.ChkHUA.Size = new System.Drawing.Size(57, 21);
            this.ChkHUA.TabIndex = 13;
            this.ChkHUA.Text = "HUA";
            this.ChkHUA.UseVisualStyleBackColor = false;
            // 
            // groupBox5
            // 
            this.groupBox5.BackColor = System.Drawing.SystemColors.Control;
            this.groupBox5.Controls.Add(this.optBadMaterial);
            this.groupBox5.Controls.Add(this.optGoodMaterial);
            this.groupBox5.Location = new System.Drawing.Point(687, 46);
            this.groupBox5.Margin = new System.Windows.Forms.Padding(2);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Padding = new System.Windows.Forms.Padding(2);
            this.groupBox5.Size = new System.Drawing.Size(202, 44);
            this.groupBox5.TabIndex = 12;
            this.groupBox5.TabStop = false;
            // 
            // optBadMaterial
            // 
            this.optBadMaterial.BackColor = System.Drawing.SystemColors.Info;
            this.optBadMaterial.Location = new System.Drawing.Point(100, 9);
            this.optBadMaterial.Margin = new System.Windows.Forms.Padding(2);
            this.optBadMaterial.Name = "optBadMaterial";
            this.optBadMaterial.Size = new System.Drawing.Size(62, 25);
            this.optBadMaterial.TabIndex = 11;
            this.optBadMaterial.Text = "Bad";
            this.optBadMaterial.UseVisualStyleBackColor = false;
            // 
            // optGoodMaterial
            // 
            this.optGoodMaterial.BackColor = System.Drawing.SystemColors.Info;
            this.optGoodMaterial.Checked = true;
            this.optGoodMaterial.Location = new System.Drawing.Point(10, 9);
            this.optGoodMaterial.Margin = new System.Windows.Forms.Padding(2);
            this.optGoodMaterial.Name = "optGoodMaterial";
            this.optGoodMaterial.Size = new System.Drawing.Size(63, 25);
            this.optGoodMaterial.TabIndex = 10;
            this.optGoodMaterial.TabStop = true;
            this.optGoodMaterial.Text = "Good";
            this.optGoodMaterial.UseVisualStyleBackColor = false;
            // 
            // txtLotCode
            // 
            this.txtLotCode.Font = new System.Drawing.Font("宋体", 11F);
            this.txtLotCode.Location = new System.Drawing.Point(421, 42);
            this.txtLotCode.Margin = new System.Windows.Forms.Padding(2);
            this.txtLotCode.Name = "txtLotCode";
            this.txtLotCode.Size = new System.Drawing.Size(174, 24);
            this.txtLotCode.TabIndex = 9;
            this.txtLotCode.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtLotCode_KeyPress);
            // 
            // txtVendorCode
            // 
            this.txtVendorCode.Font = new System.Drawing.Font("宋体", 11F);
            this.txtVendorCode.Location = new System.Drawing.Point(421, 14);
            this.txtVendorCode.Margin = new System.Windows.Forms.Padding(2);
            this.txtVendorCode.Name = "txtVendorCode";
            this.txtVendorCode.Size = new System.Drawing.Size(174, 24);
            this.txtVendorCode.TabIndex = 8;
            this.txtVendorCode.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtVendorCode_KeyPress);
            // 
            // txtQty
            // 
            this.txtQty.Font = new System.Drawing.Font("宋体", 11F);
            this.txtQty.Location = new System.Drawing.Point(122, 76);
            this.txtQty.Margin = new System.Windows.Forms.Padding(2);
            this.txtQty.Name = "txtQty";
            this.txtQty.Size = new System.Drawing.Size(159, 24);
            this.txtQty.TabIndex = 7;
            this.txtQty.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtQty_KeyPress);
            // 
            // txtDateCode
            // 
            this.txtDateCode.Font = new System.Drawing.Font("宋体", 11F);
            this.txtDateCode.Location = new System.Drawing.Point(122, 44);
            this.txtDateCode.Margin = new System.Windows.Forms.Padding(2);
            this.txtDateCode.Name = "txtDateCode";
            this.txtDateCode.Size = new System.Drawing.Size(159, 24);
            this.txtDateCode.TabIndex = 6;
            this.txtDateCode.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtDateCode_KeyPress);
            // 
            // txtCompPN
            // 
            this.txtCompPN.Font = new System.Drawing.Font("宋体", 11F);
            this.txtCompPN.Location = new System.Drawing.Point(122, 15);
            this.txtCompPN.Margin = new System.Windows.Forms.Padding(2);
            this.txtCompPN.Name = "txtCompPN";
            this.txtCompPN.Size = new System.Drawing.Size(159, 24);
            this.txtCompPN.TabIndex = 5;
            this.txtCompPN.TextChanged += new System.EventHandler(this.txtCompPN_TextChanged);
            this.txtCompPN.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtCompPN_KeyPress);
            // 
            // label7
            // 
            this.label7.BackColor = System.Drawing.SystemColors.Info;
            this.label7.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label7.Location = new System.Drawing.Point(298, 43);
            this.label7.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(119, 20);
            this.label7.TabIndex = 4;
            this.label7.Text = "LotCode";
            // 
            // label6
            // 
            this.label6.BackColor = System.Drawing.SystemColors.Info;
            this.label6.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label6.Location = new System.Drawing.Point(298, 16);
            this.label6.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(119, 20);
            this.label6.TabIndex = 3;
            this.label6.Text = "VenderCode";
            // 
            // label5
            // 
            this.label5.BackColor = System.Drawing.SystemColors.Info;
            this.label5.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.Location = new System.Drawing.Point(12, 76);
            this.label5.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(107, 20);
            this.label5.TabIndex = 2;
            this.label5.Text = "Qty";
            // 
            // label4
            // 
            this.label4.BackColor = System.Drawing.SystemColors.Info;
            this.label4.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label4.Location = new System.Drawing.Point(11, 46);
            this.label4.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(107, 20);
            this.label4.TabIndex = 1;
            this.label4.Text = "DateCode";
            // 
            // label3
            // 
            this.label3.BackColor = System.Drawing.SystemColors.Info;
            this.label3.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label3.Location = new System.Drawing.Point(11, 15);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(107, 20);
            this.label3.TabIndex = 0;
            this.label3.Text = "CompPN";
            // 
            // gridReturnComp
            // 
            this.gridReturnComp.BackgroundColor = System.Drawing.SystemColors.Control;
            this.gridReturnComp.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.gridReturnComp.Location = new System.Drawing.Point(11, 215);
            this.gridReturnComp.Margin = new System.Windows.Forms.Padding(2);
            this.gridReturnComp.Name = "gridReturnComp";
            this.gridReturnComp.RowTemplate.Height = 37;
            this.gridReturnComp.Size = new System.Drawing.Size(903, 193);
            this.gridReturnComp.TabIndex = 9;
            // 
            // LblMessage
            // 
            this.LblMessage.BackColor = System.Drawing.Color.PeachPuff;
            this.LblMessage.Font = new System.Drawing.Font("宋体", 12F);
            this.LblMessage.Location = new System.Drawing.Point(14, 421);
            this.LblMessage.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.LblMessage.Name = "LblMessage";
            this.LblMessage.Size = new System.Drawing.Size(900, 45);
            this.LblMessage.TabIndex = 10;
            this.LblMessage.Text = "                                                                                 " +
    "      ";
            // 
            // lblFeedBack
            // 
            this.lblFeedBack.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.lblFeedBack.Font = new System.Drawing.Font("宋体", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lblFeedBack.ForeColor = System.Drawing.Color.Magenta;
            this.lblFeedBack.Location = new System.Drawing.Point(13, 484);
            this.lblFeedBack.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblFeedBack.Name = "lblFeedBack";
            this.lblFeedBack.Size = new System.Drawing.Size(904, 43);
            this.lblFeedBack.TabIndex = 11;
            this.lblFeedBack.Text = "Qty FeedBack:                                                                    " +
    "           ";
            // 
            // FrmReturnComp
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(925, 549);
            this.Controls.Add(this.lblFeedBack);
            this.Controls.Add(this.LblMessage);
            this.Controls.Add(this.gridReturnComp);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Margin = new System.Windows.Forms.Padding(2);
            this.Name = "FrmReturnComp";
            this.Text = "FrmReturnComp";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FrmReturnComp_FormClosed);
            this.Load += new System.EventHandler(this.FrmReturnComp_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.groupBox5.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.gridReturnComp)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton OptSATO;
        private System.Windows.Forms.RadioButton OptZebra;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.RadioButton optNetwork;
        private System.Windows.Forms.RadioButton OptPrint;
        private System.Windows.Forms.RadioButton OptComp;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtCompPort;
        private System.Windows.Forms.TextBox txtComm;
        private System.Windows.Forms.Button CmdCommSave;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.Button cmdGetRefID;
        private System.Windows.Forms.Button cmdReprint;
        private System.Windows.Forms.Button cmdOK;
        private System.Windows.Forms.CheckBox ChkBGA;
        private System.Windows.Forms.CheckBox ChkHUA;
        private System.Windows.Forms.GroupBox groupBox5;
        private System.Windows.Forms.RadioButton optBadMaterial;
        private System.Windows.Forms.RadioButton optGoodMaterial;
        private System.Windows.Forms.TextBox txtLotCode;
        private System.Windows.Forms.TextBox txtVendorCode;
        private System.Windows.Forms.TextBox txtQty;
        private System.Windows.Forms.TextBox txtDateCode;
        private System.Windows.Forms.TextBox txtCompPN;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DataGridView gridReturnComp;
        private System.Windows.Forms.Label LblMessage;
        private System.Windows.Forms.Label lblFeedBack;
    }
}