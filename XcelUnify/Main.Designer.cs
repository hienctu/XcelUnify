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
            SuspendLayout();
            // 
            // btnGenerate
            // 
            btnGenerate.Location = new Point(12, 30);
            btnGenerate.Name = "btnGenerate";
            btnGenerate.Size = new Size(75, 23);
            btnGenerate.TabIndex = 0;
            btnGenerate.Text = "Generate";
            btnGenerate.UseVisualStyleBackColor = true;
            btnGenerate.Click += btnGenerate_Click;
            // 
            // UnifyBtn
            // 
            UnifyBtn.Location = new Point(12, 74);
            UnifyBtn.Name = "UnifyBtn";
            UnifyBtn.Size = new Size(75, 23);
            UnifyBtn.TabIndex = 1;
            UnifyBtn.Text = "Unify";
            UnifyBtn.UseVisualStyleBackColor = true;
            UnifyBtn.Click += UnifyBtn_Click;
            // 
            // Main
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(728, 427);
            Controls.Add(UnifyBtn);
            Controls.Add(btnGenerate);
            Name = "Main";
            Text = "Excel Generator & Unify Tool";
            ResumeLayout(false);
        }

        #endregion

        private Button btnGenerate;
        private Button UnifyBtn;
    }
}
