namespace ITVolume
{
    partial class ChargeByStep
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
            this.ITVolume = new System.Windows.Forms.Button();
            this.msg = new System.Windows.Forms.TextBox();
            this.Test = new System.Windows.Forms.Button();
            this.ParaTime = new System.Windows.Forms.TextBox();
            this.chargemonth = new System.Windows.Forms.Label();
            this.IT_fault_data = new System.Windows.Forms.Button();
            this.Voice_fault_data = new System.Windows.Forms.Button();
            this.double_charge = new System.Windows.Forms.Button();
            this.atos = new System.Windows.Forms.Button();
            this.FileStorePath = new System.Windows.Forms.Label();
            this.ExcelStorePath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.directline = new System.Windows.Forms.Button();
            this.newcc = new System.Windows.Forms.Button();
            this.Snx = new System.Windows.Forms.Button();
            this.SnxTempInvalidCC = new System.Windows.Forms.Button();
            this.SnxInvalidCC = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.ImportMatrix = new System.Windows.Forms.Button();
            this.HR = new System.Windows.Forms.Button();
            this.CIT_with_GID = new System.Windows.Forms.Button();
            this.CIT_with_EMAIL = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.generate_hr_txt = new System.Windows.Forms.Button();
            this.Import_CIT_flender = new System.Windows.Forms.Button();
            this.LoaclDB = new System.Windows.Forms.CheckBox();
            this.WCWUserList = new System.Windows.Forms.Button();
            this.WcwChargeScope = new System.Windows.Forms.Button();
            this.WCWCCCheck = new System.Windows.Forms.Button();
            this.CheckExcelCC = new System.Windows.Forms.Button();
            this.ChangeInitConfig = new System.Windows.Forms.Button();
            this.CD_Report_New = new System.Windows.Forms.Button();
            this.m_bgWorker = new System.ComponentModel.BackgroundWorker();
            this.MainProgressBar = new System.Windows.Forms.ProgressBar();
            this.CheckBUChanged = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // ITVolume
            // 
            this.ITVolume.Location = new System.Drawing.Point(561, 44);
            this.ITVolume.Name = "ITVolume";
            this.ITVolume.Size = new System.Drawing.Size(163, 23);
            this.ITVolume.TabIndex = 0;
            this.ITVolume.Text = "CreateMonthlyOSMQVolume";
            this.ITVolume.UseVisualStyleBackColor = true;
            this.ITVolume.Click += new System.EventHandler(this.ITVolume_Click);
            // 
            // msg
            // 
            this.msg.Location = new System.Drawing.Point(12, 352);
            this.msg.Multiline = true;
            this.msg.Name = "msg";
            this.msg.Size = new System.Drawing.Size(847, 189);
            this.msg.TabIndex = 1;
            this.msg.Text = "message";
            // 
            // Test
            // 
            this.Test.Location = new System.Drawing.Point(107, 317);
            this.Test.Name = "Test";
            this.Test.Size = new System.Drawing.Size(75, 23);
            this.Test.TabIndex = 2;
            this.Test.Text = "Test";
            this.Test.UseVisualStyleBackColor = true;
            this.Test.Click += new System.EventHandler(this.Test_Click);
            // 
            // ParaTime
            // 
            this.ParaTime.Location = new System.Drawing.Point(12, 12);
            this.ParaTime.Name = "ParaTime";
            this.ParaTime.Size = new System.Drawing.Size(107, 20);
            this.ParaTime.TabIndex = 3;
            this.ParaTime.TextChanged += new System.EventHandler(this.ParaTime_TextChanged);
            // 
            // chargemonth
            // 
            this.chargemonth.AutoSize = true;
            this.chargemonth.Location = new System.Drawing.Point(13, -1);
            this.chargemonth.Name = "chargemonth";
            this.chargemonth.Size = new System.Drawing.Size(69, 13);
            this.chargemonth.TabIndex = 4;
            this.chargemonth.Text = "chargemonth";
            // 
            // IT_fault_data
            // 
            this.IT_fault_data.Location = new System.Drawing.Point(457, 15);
            this.IT_fault_data.Name = "IT_fault_data";
            this.IT_fault_data.Size = new System.Drawing.Size(88, 23);
            this.IT_fault_data.TabIndex = 5;
            this.IT_fault_data.Text = "IT fault data";
            this.IT_fault_data.UseVisualStyleBackColor = true;
            this.IT_fault_data.Click += new System.EventHandler(this.IT_fault_data_Click);
            // 
            // Voice_fault_data
            // 
            this.Voice_fault_data.Location = new System.Drawing.Point(561, 15);
            this.Voice_fault_data.Name = "Voice_fault_data";
            this.Voice_fault_data.Size = new System.Drawing.Size(98, 23);
            this.Voice_fault_data.TabIndex = 6;
            this.Voice_fault_data.Text = "Voice fault data";
            this.Voice_fault_data.UseVisualStyleBackColor = true;
            this.Voice_fault_data.Click += new System.EventHandler(this.Voice_fault_data_Click);
            // 
            // double_charge
            // 
            this.double_charge.Location = new System.Drawing.Point(675, 15);
            this.double_charge.Name = "double_charge";
            this.double_charge.Size = new System.Drawing.Size(92, 23);
            this.double_charge.TabIndex = 7;
            this.double_charge.Text = "Double Charge";
            this.double_charge.UseVisualStyleBackColor = true;
            this.double_charge.Click += new System.EventHandler(this.double_charge_Click);
            // 
            // atos
            // 
            this.atos.Location = new System.Drawing.Point(784, 15);
            this.atos.Name = "atos";
            this.atos.Size = new System.Drawing.Size(75, 23);
            this.atos.TabIndex = 8;
            this.atos.Text = "Atos";
            this.atos.UseVisualStyleBackColor = true;
            this.atos.Click += new System.EventHandler(this.atos_Click);
            // 
            // FileStorePath
            // 
            this.FileStorePath.AutoSize = true;
            this.FileStorePath.Location = new System.Drawing.Point(13, 35);
            this.FileStorePath.Name = "FileStorePath";
            this.FileStorePath.Size = new System.Drawing.Size(70, 13);
            this.FileStorePath.TabIndex = 9;
            this.FileStorePath.Text = "FileStorePath";
            // 
            // ExcelStorePath
            // 
            this.ExcelStorePath.Location = new System.Drawing.Point(12, 51);
            this.ExcelStorePath.Name = "ExcelStorePath";
            this.ExcelStorePath.Size = new System.Drawing.Size(192, 20);
            this.ExcelStorePath.TabIndex = 10;
            this.ExcelStorePath.TextChanged += new System.EventHandler(this.ExcelStorePath_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(454, -1);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 13);
            this.label1.TabIndex = 11;
            this.label1.Text = "ReportButton";
            // 
            // directline
            // 
            this.directline.Location = new System.Drawing.Point(12, 106);
            this.directline.Name = "directline";
            this.directline.Size = new System.Drawing.Size(162, 23);
            this.directline.TabIndex = 12;
            this.directline.Text = "directline";
            this.directline.UseVisualStyleBackColor = true;
            this.directline.Click += new System.EventHandler(this.directline_Click);
            // 
            // newcc
            // 
            this.newcc.Location = new System.Drawing.Point(457, 44);
            this.newcc.Name = "newcc";
            this.newcc.Size = new System.Drawing.Size(75, 23);
            this.newcc.TabIndex = 13;
            this.newcc.Text = "NewCC";
            this.newcc.UseVisualStyleBackColor = true;
            this.newcc.Click += new System.EventHandler(this.newcc_Click);
            // 
            // Snx
            // 
            this.Snx.Location = new System.Drawing.Point(99, 149);
            this.Snx.Name = "Snx";
            this.Snx.Size = new System.Drawing.Size(75, 23);
            this.Snx.TabIndex = 14;
            this.Snx.Text = "ImportSnx";
            this.Snx.UseVisualStyleBackColor = true;
            this.Snx.Click += new System.EventHandler(this.Snx_Click);
            // 
            // SnxTempInvalidCC
            // 
            this.SnxTempInvalidCC.Location = new System.Drawing.Point(180, 149);
            this.SnxTempInvalidCC.Name = "SnxTempInvalidCC";
            this.SnxTempInvalidCC.Size = new System.Drawing.Size(116, 23);
            this.SnxTempInvalidCC.TabIndex = 15;
            this.SnxTempInvalidCC.Text = "SnxTempInvalidCC";
            this.SnxTempInvalidCC.UseVisualStyleBackColor = true;
            this.SnxTempInvalidCC.Click += new System.EventHandler(this.SnxTempInvalidCC_Click);
            // 
            // SnxInvalidCC
            // 
            this.SnxInvalidCC.Location = new System.Drawing.Point(302, 149);
            this.SnxInvalidCC.Name = "SnxInvalidCC";
            this.SnxInvalidCC.Size = new System.Drawing.Size(82, 23);
            this.SnxInvalidCC.TabIndex = 16;
            this.SnxInvalidCC.Text = "SnxInvalidCC";
            this.SnxInvalidCC.UseVisualStyleBackColor = true;
            this.SnxInvalidCC.Click += new System.EventHandler(this.SnxInvalidCC_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 90);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(76, 13);
            this.label2.TabIndex = 17;
            this.label2.Text = "ProcessButton";
            // 
            // ImportMatrix
            // 
            this.ImportMatrix.Location = new System.Drawing.Point(12, 149);
            this.ImportMatrix.Name = "ImportMatrix";
            this.ImportMatrix.Size = new System.Drawing.Size(78, 23);
            this.ImportMatrix.TabIndex = 18;
            this.ImportMatrix.Text = "ImportMatrix";
            this.ImportMatrix.UseVisualStyleBackColor = true;
            this.ImportMatrix.Click += new System.EventHandler(this.ImportMatrix_Click);
            // 
            // HR
            // 
            this.HR.Location = new System.Drawing.Point(12, 187);
            this.HR.Name = "HR";
            this.HR.Size = new System.Drawing.Size(162, 23);
            this.HR.TabIndex = 19;
            this.HR.Text = "clean_HR_data";
            this.HR.UseVisualStyleBackColor = true;
            this.HR.Click += new System.EventHandler(this.HR_Click);
            // 
            // CIT_with_GID
            // 
            this.CIT_with_GID.Location = new System.Drawing.Point(457, 95);
            this.CIT_with_GID.Name = "CIT_with_GID";
            this.CIT_with_GID.Size = new System.Drawing.Size(355, 23);
            this.CIT_with_GID.TabIndex = 20;
            this.CIT_with_GID.Text = "Import_CIT_with_GID";
            this.CIT_with_GID.UseVisualStyleBackColor = true;
            this.CIT_with_GID.Click += new System.EventHandler(this.CIT_with_GID_Click);
            // 
            // CIT_with_EMAIL
            // 
            this.CIT_with_EMAIL.Location = new System.Drawing.Point(457, 124);
            this.CIT_with_EMAIL.Name = "CIT_with_EMAIL";
            this.CIT_with_EMAIL.Size = new System.Drawing.Size(355, 23);
            this.CIT_with_EMAIL.TabIndex = 21;
            this.CIT_with_EMAIL.Text = "Import_CIT_with_EMAIL";
            this.CIT_with_EMAIL.UseVisualStyleBackColor = true;
            this.CIT_with_EMAIL.Click += new System.EventHandler(this.CIT_with_EMAIL_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(454, 79);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(80, 13);
            this.label3.TabIndex = 22;
            this.label3.Text = "CITSeviceData";
            // 
            // generate_hr_txt
            // 
            this.generate_hr_txt.Location = new System.Drawing.Point(180, 187);
            this.generate_hr_txt.Name = "generate_hr_txt";
            this.generate_hr_txt.Size = new System.Drawing.Size(204, 23);
            this.generate_hr_txt.TabIndex = 23;
            this.generate_hr_txt.Text = "Generate Hr CMD Txt ";
            this.generate_hr_txt.UseVisualStyleBackColor = true;
            this.generate_hr_txt.Click += new System.EventHandler(this.generate_hr_txt_Click);
            // 
            // Import_CIT_flender
            // 
            this.Import_CIT_flender.Location = new System.Drawing.Point(457, 153);
            this.Import_CIT_flender.Name = "Import_CIT_flender";
            this.Import_CIT_flender.Size = new System.Drawing.Size(355, 23);
            this.Import_CIT_flender.TabIndex = 24;
            this.Import_CIT_flender.Text = "Import_CIT_flender";
            this.Import_CIT_flender.UseVisualStyleBackColor = true;
            this.Import_CIT_flender.Click += new System.EventHandler(this.Import_CIT_flender_Click);
            // 
            // LoaclDB
            // 
            this.LoaclDB.AutoSize = true;
            this.LoaclDB.Location = new System.Drawing.Point(15, 323);
            this.LoaclDB.Name = "LoaclDB";
            this.LoaclDB.Size = new System.Drawing.Size(86, 17);
            this.LoaclDB.TabIndex = 25;
            this.LoaclDB.Text = "UseLoaclDB";
            this.LoaclDB.UseVisualStyleBackColor = true;
            this.LoaclDB.CheckedChanged += new System.EventHandler(this.LoaclDB_CheckedChanged);
            // 
            // WCWUserList
            // 
            this.WCWUserList.Location = new System.Drawing.Point(632, 216);
            this.WCWUserList.Name = "WCWUserList";
            this.WCWUserList.Size = new System.Drawing.Size(180, 23);
            this.WCWUserList.TabIndex = 26;
            this.WCWUserList.Text = "WCW2 WCWUserList";
            this.WCWUserList.UseVisualStyleBackColor = true;
            // 
            // WcwChargeScope
            // 
            this.WcwChargeScope.Location = new System.Drawing.Point(632, 187);
            this.WcwChargeScope.Name = "WcwChargeScope";
            this.WcwChargeScope.Size = new System.Drawing.Size(180, 23);
            this.WcwChargeScope.TabIndex = 27;
            this.WcwChargeScope.Text = "WCW1WcwChargeScope";
            this.WcwChargeScope.UseVisualStyleBackColor = true;
            this.WcwChargeScope.Click += new System.EventHandler(this.WcwChargeScope_Click);
            // 
            // WCWCCCheck
            // 
            this.WCWCCCheck.Location = new System.Drawing.Point(632, 245);
            this.WCWCCCheck.Name = "WCWCCCheck";
            this.WCWCCCheck.Size = new System.Drawing.Size(180, 23);
            this.WCWCCCheck.TabIndex = 28;
            this.WCWCCCheck.Text = "WCW3 Check CC";
            this.WCWCCCheck.UseVisualStyleBackColor = true;
            // 
            // CheckExcelCC
            // 
            this.CheckExcelCC.Location = new System.Drawing.Point(15, 254);
            this.CheckExcelCC.Name = "CheckExcelCC";
            this.CheckExcelCC.Size = new System.Drawing.Size(139, 23);
            this.CheckExcelCC.TabIndex = 29;
            this.CheckExcelCC.Text = "CheckExcelCCData";
            this.CheckExcelCC.UseVisualStyleBackColor = true;
            this.CheckExcelCC.Click += new System.EventHandler(this.CheckExcelCC_Click);
            // 
            // ChangeInitConfig
            // 
            this.ChangeInitConfig.Location = new System.Drawing.Point(202, 12);
            this.ChangeInitConfig.Name = "ChangeInitConfig";
            this.ChangeInitConfig.Size = new System.Drawing.Size(147, 23);
            this.ChangeInitConfig.TabIndex = 31;
            this.ChangeInitConfig.Text = "ChangeInitConfig";
            this.ChangeInitConfig.UseVisualStyleBackColor = true;
            this.ChangeInitConfig.Click += new System.EventHandler(this.ChangeInitConfig_Click);
            // 
            // CD_Report_New
            // 
            this.CD_Report_New.Location = new System.Drawing.Point(160, 254);
            this.CD_Report_New.Name = "CD_Report_New";
            this.CD_Report_New.Size = new System.Drawing.Size(110, 23);
            this.CD_Report_New.TabIndex = 32;
            this.CD_Report_New.Text = "CD_Report_New";
            this.CD_Report_New.UseVisualStyleBackColor = true;
            this.CD_Report_New.Click += new System.EventHandler(this.CD_Report_New_Click);
            // 
            // m_bgWorker
            // 
            this.m_bgWorker.WorkerReportsProgress = true;
            this.m_bgWorker.WorkerSupportsCancellation = true;
            this.m_bgWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.m_bgWorker_DoWork);
            // 
            // MainProgressBar
            // 
            this.MainProgressBar.Location = new System.Drawing.Point(291, 317);
            this.MainProgressBar.Name = "MainProgressBar";
            this.MainProgressBar.Size = new System.Drawing.Size(243, 23);
            this.MainProgressBar.TabIndex = 33;
            // 
            // CheckBUChanged
            // 
            this.CheckBUChanged.Location = new System.Drawing.Point(302, 254);
            this.CheckBUChanged.Name = "CheckBUChanged";
            this.CheckBUChanged.Size = new System.Drawing.Size(140, 23);
            this.CheckBUChanged.TabIndex = 34;
            this.CheckBUChanged.Text = "CheckBUChanged";
            this.CheckBUChanged.UseVisualStyleBackColor = true;
            this.CheckBUChanged.Click += new System.EventHandler(this.CheckBUChanged_Click);
            // 
            // ChargeByStep
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(871, 553);
            this.Controls.Add(this.CheckBUChanged);
            this.Controls.Add(this.MainProgressBar);
            this.Controls.Add(this.CD_Report_New);
            this.Controls.Add(this.ChangeInitConfig);
            this.Controls.Add(this.CheckExcelCC);
            this.Controls.Add(this.WCWCCCheck);
            this.Controls.Add(this.WcwChargeScope);
            this.Controls.Add(this.WCWUserList);
            this.Controls.Add(this.LoaclDB);
            this.Controls.Add(this.Import_CIT_flender);
            this.Controls.Add(this.generate_hr_txt);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.CIT_with_EMAIL);
            this.Controls.Add(this.CIT_with_GID);
            this.Controls.Add(this.HR);
            this.Controls.Add(this.ImportMatrix);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.SnxInvalidCC);
            this.Controls.Add(this.SnxTempInvalidCC);
            this.Controls.Add(this.Snx);
            this.Controls.Add(this.newcc);
            this.Controls.Add(this.directline);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.ExcelStorePath);
            this.Controls.Add(this.FileStorePath);
            this.Controls.Add(this.atos);
            this.Controls.Add(this.double_charge);
            this.Controls.Add(this.Voice_fault_data);
            this.Controls.Add(this.IT_fault_data);
            this.Controls.Add(this.chargemonth);
            this.Controls.Add(this.ParaTime);
            this.Controls.Add(this.Test);
            this.Controls.Add(this.msg);
            this.Controls.Add(this.ITVolume);
            this.Name = "ChargeByStep";
            this.Text = "ChargeByStep";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button ITVolume;
        private System.Windows.Forms.TextBox msg;
        private System.Windows.Forms.Button Test;
        private System.Windows.Forms.TextBox ParaTime;
        private System.Windows.Forms.Label chargemonth;
        private System.Windows.Forms.Button IT_fault_data;
        private System.Windows.Forms.Button Voice_fault_data;
        private System.Windows.Forms.Button double_charge;
        private System.Windows.Forms.Button atos;
        private System.Windows.Forms.Label FileStorePath;
        private System.Windows.Forms.TextBox ExcelStorePath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button directline;
        private System.Windows.Forms.Button newcc;
        private System.Windows.Forms.Button Snx;
        private System.Windows.Forms.Button SnxTempInvalidCC;
        private System.Windows.Forms.Button SnxInvalidCC;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button ImportMatrix;
        private System.Windows.Forms.Button HR;
        private System.Windows.Forms.Button CIT_with_GID;
        private System.Windows.Forms.Button CIT_with_EMAIL;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button generate_hr_txt;
        private System.Windows.Forms.Button Import_CIT_flender;
        private System.Windows.Forms.CheckBox LoaclDB;
        private System.Windows.Forms.Button WCWUserList;
        private System.Windows.Forms.Button WcwChargeScope;
        private System.Windows.Forms.Button WCWCCCheck;
        private System.Windows.Forms.Button CheckExcelCC;
        private System.Windows.Forms.Button ChangeInitConfig;
        private System.Windows.Forms.Button CD_Report_New;
        private System.ComponentModel.BackgroundWorker m_bgWorker;
        private System.Windows.Forms.ProgressBar MainProgressBar;
        private System.Windows.Forms.Button CheckBUChanged;
        //private System.Windows.Forms.Button Initfloder;
    }
}

