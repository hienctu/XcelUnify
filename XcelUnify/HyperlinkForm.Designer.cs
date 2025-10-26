namespace XcelUnify
{
    partial class HyperlinkForm
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
            btnYes = new Button();
            btnNo = new Button();
            label1 = new Label();
            label2 = new Label();
            hplStaffUpdateTempFolder = new LinkLabel();
            hplSharePointOutputLocation = new LinkLabel();
            label3 = new Label();
            SuspendLayout();
            // 
            // btnYes
            // 
            btnYes.DialogResult = DialogResult.Yes;
            btnYes.Location = new Point(413, 155);
            btnYes.Name = "btnYes";
            btnYes.Size = new Size(75, 23);
            btnYes.TabIndex = 0;
            btnYes.Text = "Yes";
            btnYes.UseVisualStyleBackColor = true;
            // 
            // btnNo
            // 
            btnNo.DialogResult = DialogResult.No;
            btnNo.Location = new Point(332, 155);
            btnNo.Name = "btnNo";
            btnNo.Size = new Size(75, 23);
            btnNo.TabIndex = 1;
            btnNo.Text = "No";
            btnNo.UseVisualStyleBackColor = true;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(14, 84);
            label1.Name = "label1";
            label1.Size = new Size(144, 15);
            label1.TabIndex = 2;
            label1.Text = "Staff Update Temp Folder:";
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Location = new Point(14, 120);
            label2.Name = "label2";
            label2.Size = new Size(157, 15);
            label2.TabIndex = 3;
            label2.Text = "SharePoint Output Location:";
            // 
            // hplStaffUpdateTempFolder
            // 
            hplStaffUpdateTempFolder.AutoSize = true;
            hplStaffUpdateTempFolder.Location = new Point(208, 84);
            hplStaffUpdateTempFolder.Name = "hplStaffUpdateTempFolder";
            hplStaffUpdateTempFolder.Size = new Size(60, 15);
            hplStaffUpdateTempFolder.TabIndex = 4;
            hplStaffUpdateTempFolder.TabStop = true;
            hplStaffUpdateTempFolder.Text = "linkLabel1";
            // 
            // hplSharePointOutputLocation
            // 
            hplSharePointOutputLocation.AutoSize = true;
            hplSharePointOutputLocation.Location = new Point(208, 120);
            hplSharePointOutputLocation.Name = "hplSharePointOutputLocation";
            hplSharePointOutputLocation.Size = new Size(60, 15);
            hplSharePointOutputLocation.TabIndex = 5;
            hplSharePointOutputLocation.TabStop = true;
            hplSharePointOutputLocation.Text = "linkLabel2";
            // 
            // label3
            // 
            label3.Location = new Point(14, 22);
            label3.Name = "label3";
            label3.Size = new Size(471, 51);
            label3.TabIndex = 6;
            label3.Text = "Are you sure you want to upload all files in Staff Update Temp Folder overwrite existing files in SharePoint Output Location?";
            // 
            // HyperlinkForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(501, 190);
            Controls.Add(label3);
            Controls.Add(hplSharePointOutputLocation);
            Controls.Add(hplStaffUpdateTempFolder);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(btnNo);
            Controls.Add(btnYes);
            Name = "HyperlinkForm";
            StartPosition = FormStartPosition.CenterParent;
            Text = "HyperlinkForm";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnYes;
        private Button btnNo;
        private Label label1;
        private Label label2;
        private LinkLabel hplStaffUpdateTempFolder;
        private LinkLabel hplSharePointOutputLocation;
        private Label label3;
    }
}