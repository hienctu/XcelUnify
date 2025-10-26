using System.Diagnostics;

namespace XcelUnify
{
    public partial class HyperlinkForm : Form
    {
        public HyperlinkForm()
        {
            InitializeComponent();
        }

        public HyperlinkForm(string tempFolder, string sharepointFolder)
        {
            InitializeComponent();

            // Create hyperlinks for the specified folders
            CreateHyperlink(hplStaffUpdateTempFolder, tempFolder);
            CreateHyperlink(hplSharePointOutputLocation, sharepointFolder);
        }

        private void CreateHyperlink(Label label, string folderPath)
        {
            label.Text = "View Folder";
            label.ForeColor = Color.Blue;
            label.Cursor = Cursors.Hand;
            label.Click += (sender, e) => OpenFolder(folderPath);
        }

        private void OpenFolder(string folderPath)
        {
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = folderPath,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening folder: " + ex.Message);
            }
        }
    }
}
