namespace XcelUnify
{
    partial class Main
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            btnGenerate = new Button();
            UnifyBtn = new Button();
            lblApplicationTitle = new Label();
            lblMasterFile = new Label();
            groupBox1 = new GroupBox();
            btnViewDualCampusTemplate = new Button();
            txtDualCampusTemplateFile = new TextBox();
            label3 = new Label();
            btnViewResearchTemplate = new Button();
            txtResearchTemplateFile = new TextBox();
            label2 = new Label();
            lblMasterFileRowCount = new Label();
            btnViewTemplate = new Button();
            btnViewMaster = new Button();
            txtTemplateFile = new TextBox();
            txtMasterFile = new TextBox();
            label1 = new Label();
            groupBox2 = new GroupBox();
            btnClose = new Button();
            btnCloseExcels = new Button();
            progressBar = new ProgressBar();
            lblActionDisplay = new Label();
            lstReport = new ListBox();
            lblReport = new Label();
            btnViewOutput = new Button();
            groupBox1.SuspendLayout();
            groupBox2.SuspendLayout();
            SuspendLayout();
            // 
            // btnGenerate
            // 
            btnGenerate.AllowDrop = true;
            btnGenerate.Location = new Point(13, 76);
            btnGenerate.Name = "btnGenerate";
            btnGenerate.Size = new Size(113, 40);
            btnGenerate.TabIndex = 0;
            btnGenerate.Text = "Generate SAFES workloads";
            btnGenerate.UseVisualStyleBackColor = true;
            btnGenerate.Click += btnGenerate_Click;
            // 
            // UnifyBtn
            // 
            UnifyBtn.Location = new Point(13, 134);
            UnifyBtn.Name = "UnifyBtn";
            UnifyBtn.Size = new Size(113, 34);
            UnifyBtn.TabIndex = 1;
            UnifyBtn.Text = "Unify workloads";
            UnifyBtn.UseVisualStyleBackColor = true;
            UnifyBtn.Click += UnifyBtn_Click;
            // 
            // lblApplicationTitle
            // 
            lblApplicationTitle.AutoSize = true;
            lblApplicationTitle.Font = new Font("Lucida Bright", 15.75F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lblApplicationTitle.Location = new Point(12, 20);
            lblApplicationTitle.Name = "lblApplicationTitle";
            lblApplicationTitle.Size = new Size(304, 24);
            lblApplicationTitle.TabIndex = 2;
            lblApplicationTitle.Text = "SAFES Workload Generator";
            // 
            // lblMasterFile
            // 
            lblMasterFile.AutoSize = true;
            lblMasterFile.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0);
            lblMasterFile.Location = new Point(6, 19);
            lblMasterFile.Name = "lblMasterFile";
            lblMasterFile.Size = new Size(113, 17);
            lblMasterFile.TabIndex = 3;
            lblMasterFile.Text = "Master Data File:";
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(btnViewDualCampusTemplate);
            groupBox1.Controls.Add(txtDualCampusTemplateFile);
            groupBox1.Controls.Add(label3);
            groupBox1.Controls.Add(btnViewResearchTemplate);
            groupBox1.Controls.Add(txtResearchTemplateFile);
            groupBox1.Controls.Add(label2);
            groupBox1.Controls.Add(lblMasterFileRowCount);
            groupBox1.Controls.Add(btnViewTemplate);
            groupBox1.Controls.Add(btnViewMaster);
            groupBox1.Controls.Add(txtTemplateFile);
            groupBox1.Controls.Add(txtMasterFile);
            groupBox1.Controls.Add(label1);
            groupBox1.Controls.Add(lblMasterFile);
            groupBox1.Location = new Point(16, 57);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(700, 232);
            groupBox1.TabIndex = 4;
            groupBox1.TabStop = false;
            // 
            // btnViewDualCampusTemplate
            // 
            btnViewDualCampusTemplate.Location = new Point(601, 180);
            btnViewDualCampusTemplate.Name = "btnViewDualCampusTemplate";
            btnViewDualCampusTemplate.Size = new Size(75, 23);
            btnViewDualCampusTemplate.TabIndex = 15;
            btnViewDualCampusTemplate.Text = "View File";
            btnViewDualCampusTemplate.UseVisualStyleBackColor = true;
            btnViewDualCampusTemplate.Click += btnViewDualCampusTemplate_Click;
            // 
            // txtDualCampusTemplateFile
            // 
            txtDualCampusTemplateFile.Location = new Point(224, 180);
            txtDualCampusTemplateFile.Name = "txtDualCampusTemplateFile";
            txtDualCampusTemplateFile.ReadOnly = true;
            txtDualCampusTemplateFile.Size = new Size(357, 23);
            txtDualCampusTemplateFile.TabIndex = 14;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label3.Location = new Point(7, 183);
            label3.Name = "label3";
            label3.Size = new Size(155, 17);
            label3.TabIndex = 13;
            label3.Text = "Dual Campus Template:";
            // 
            // btnViewResearchTemplate
            // 
            btnViewResearchTemplate.Location = new Point(601, 129);
            btnViewResearchTemplate.Name = "btnViewResearchTemplate";
            btnViewResearchTemplate.Size = new Size(75, 23);
            btnViewResearchTemplate.TabIndex = 12;
            btnViewResearchTemplate.Text = "View File";
            btnViewResearchTemplate.UseVisualStyleBackColor = true;
            btnViewResearchTemplate.Click += button1_Click;
            // 
            // txtResearchTemplateFile
            // 
            txtResearchTemplateFile.Location = new Point(224, 129);
            txtResearchTemplateFile.Name = "txtResearchTemplateFile";
            txtResearchTemplateFile.ReadOnly = true;
            txtResearchTemplateFile.Size = new Size(357, 23);
            txtResearchTemplateFile.TabIndex = 11;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label2.Location = new Point(7, 132);
            label2.Name = "label2";
            label2.Size = new Size(196, 17);
            label2.TabIndex = 10;
            label2.Text = "Research/Internship Template:";
            // 
            // lblMasterFileRowCount
            // 
            lblMasterFileRowCount.AutoSize = true;
            lblMasterFileRowCount.Location = new Point(224, 44);
            lblMasterFileRowCount.Name = "lblMasterFileRowCount";
            lblMasterFileRowCount.Size = new Size(38, 15);
            lblMasterFileRowCount.TabIndex = 9;
            lblMasterFileRowCount.Text = "label2";
            // 
            // btnViewTemplate
            // 
            btnViewTemplate.Location = new Point(601, 75);
            btnViewTemplate.Name = "btnViewTemplate";
            btnViewTemplate.Size = new Size(75, 23);
            btnViewTemplate.TabIndex = 8;
            btnViewTemplate.Text = "View File";
            btnViewTemplate.UseVisualStyleBackColor = true;
            btnViewTemplate.Click += btnViewTemplate_Click;
            // 
            // btnViewMaster
            // 
            btnViewMaster.Location = new Point(601, 18);
            btnViewMaster.Name = "btnViewMaster";
            btnViewMaster.Size = new Size(75, 23);
            btnViewMaster.TabIndex = 7;
            btnViewMaster.Text = "View File";
            btnViewMaster.UseVisualStyleBackColor = true;
            btnViewMaster.Click += btnViewMaster_Click;
            // 
            // txtTemplateFile
            // 
            txtTemplateFile.Location = new Point(224, 74);
            txtTemplateFile.Name = "txtTemplateFile";
            txtTemplateFile.ReadOnly = true;
            txtTemplateFile.Size = new Size(357, 23);
            txtTemplateFile.TabIndex = 6;
            // 
            // txtMasterFile
            // 
            txtMasterFile.Location = new Point(224, 18);
            txtMasterFile.Name = "txtMasterFile";
            txtMasterFile.ReadOnly = true;
            txtMasterFile.Size = new Size(357, 23);
            txtMasterFile.TabIndex = 5;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI", 9.75F, FontStyle.Bold, GraphicsUnit.Point, 0);
            label1.Location = new Point(6, 77);
            label1.Name = "label1";
            label1.Size = new Size(128, 17);
            label1.TabIndex = 4;
            label1.Text = "Standard Template:";
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(btnClose);
            groupBox2.Controls.Add(btnCloseExcels);
            groupBox2.Controls.Add(UnifyBtn);
            groupBox2.Controls.Add(btnGenerate);
            groupBox2.Location = new Point(16, 304);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(149, 315);
            groupBox2.TabIndex = 5;
            groupBox2.TabStop = false;
            groupBox2.Text = "Action";
            // 
            // btnClose
            // 
            btnClose.Location = new Point(13, 273);
            btnClose.Name = "btnClose";
            btnClose.Size = new Size(113, 23);
            btnClose.TabIndex = 3;
            btnClose.Text = "Close Form";
            btnClose.UseVisualStyleBackColor = true;
            btnClose.Click += btnClose_Click;
            // 
            // btnCloseExcels
            // 
            btnCloseExcels.Location = new Point(13, 33);
            btnCloseExcels.Name = "btnCloseExcels";
            btnCloseExcels.Size = new Size(113, 23);
            btnCloseExcels.TabIndex = 2;
            btnCloseExcels.Text = "Close All Excels";
            btnCloseExcels.UseVisualStyleBackColor = true;
            btnCloseExcels.Click += btnCloseExcels_Click;
            // 
            // progressBar
            // 
            progressBar.Location = new Point(175, 337);
            progressBar.Name = "progressBar";
            progressBar.Size = new Size(541, 23);
            progressBar.TabIndex = 6;
            // 
            // lblActionDisplay
            // 
            lblActionDisplay.AutoSize = true;
            lblActionDisplay.Location = new Point(175, 313);
            lblActionDisplay.Name = "lblActionDisplay";
            lblActionDisplay.Size = new Size(0, 15);
            lblActionDisplay.TabIndex = 7;
            // 
            // lstReport
            // 
            lstReport.FormattingEnabled = true;
            lstReport.ItemHeight = 15;
            lstReport.Location = new Point(175, 405);
            lstReport.Name = "lstReport";
            lstReport.Size = new Size(541, 214);
            lstReport.TabIndex = 8;
            // 
            // lblReport
            // 
            lblReport.AutoSize = true;
            lblReport.Location = new Point(175, 380);
            lblReport.Name = "lblReport";
            lblReport.Size = new Size(65, 15);
            lblReport.TabIndex = 9;
            lblReport.Text = "Generating";
            // 
            // btnViewOutput
            // 
            btnViewOutput.Location = new Point(570, 380);
            btnViewOutput.Name = "btnViewOutput";
            btnViewOutput.Size = new Size(146, 23);
            btnViewOutput.TabIndex = 10;
            btnViewOutput.Text = "btnViewOutput";
            btnViewOutput.UseVisualStyleBackColor = true;
            btnViewOutput.Visible = false;
            btnViewOutput.Click += btnViewOutput_Click;
            // 
            // Main
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(728, 638);
            Controls.Add(btnViewOutput);
            Controls.Add(lblReport);
            Controls.Add(lstReport);
            Controls.Add(lblActionDisplay);
            Controls.Add(progressBar);
            Controls.Add(groupBox2);
            Controls.Add(groupBox1);
            Controls.Add(lblApplicationTitle);
            Name = "Main";
            Text = "Excel Generator & Unify Tool";
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            groupBox2.ResumeLayout(false);
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnGenerate;
        private Button UnifyBtn;
        private Label lblApplicationTitle;
        private Label lblMasterFile;
        private GroupBox groupBox1;
        private Label label1;
        private TextBox txtTemplateFile;
        private TextBox txtMasterFile;
        private Button btnViewTemplate;
        private Button btnViewMaster;
        private GroupBox groupBox2;
        private Button btnCloseExcels;
        private Label lblMasterFileRowCount;
        private ProgressBar progressBar;
        private Label lblActionDisplay;
        private ListBox lstReport;
        private Label lblReport;
        private Button btnViewOutput;
        private Button btnClose;
        private Button btnViewDualCampusTemplate;
        private TextBox txtDualCampusTemplateFile;
        private Label label3;
        private Button btnViewResearchTemplate;
        private TextBox txtResearchTemplateFile;
        private Label label2;
    }
}
