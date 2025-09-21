using Microsoft.Office.Interop.Excel;
using XcelUnify.Helpers;
using Range = Microsoft.Office.Interop.Excel.Range;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Diagnostics;

namespace XcelUnify
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
            txtMasterFile.Text = ConfigManager.Master_File;
            txtTemplateFile.ReadOnly = true;
            txtTemplateFile.Text = ConfigManager.Template_File;
            txtTemplateFile.ReadOnly = true;

            lblActionDisplay.Visible = false;
            progressBar.Visible = false;
            lblReport.Visible = false;
            lstReport.Visible = false;


            var toolTips = new ToolTip
            {
                AutoPopDelay = 5000, // Time in milliseconds the tooltip remains visible
                InitialDelay = 500,  // Delay before the tooltip appears
                ReshowDelay = 200,   // Delay before reappearing after hiding
                ShowAlways = true    // Ensures the tooltip shows even if the form is inactive
            };
            toolTips.SetToolTip(btnCloseExcels, "Close all currently running Excel processes (recommended before starting data generation or collection)");

        }

        private async void btnGenerate_Click(object sender, EventArgs e)
        {
            // Change the cursor to "Wait"
            Cursor = Cursors.WaitCursor;

            Invoke(new System.Action(() =>
            {
                lblActionDisplay.Visible = true;
                lblActionDisplay.Text = "Generating SAFES workload files...";
                progressBar.Visible = true;
                progressBar.Style = ProgressBarStyle.Marquee;
            }));

            int maxRows = ConfigManager.Max_Rows;
            int rowCount = 0;
            int colCount = 0;

            var masterFilePath = ConfigManager.Master_File;
            var templateFilePath = ConfigManager.Template_File;

            Application excelApp = null;
            Workbook masterFile = null;
            Worksheet worksheet = null;
            Range usedRange = null;

            try
            {
                excelApp = new Application();
                masterFile = excelApp.Workbooks.Open(masterFilePath);
                worksheet = (Worksheet?)masterFile.Worksheets[1];
                usedRange = worksheet.UsedRange;
                rowCount = usedRange.Rows.Count;
                colCount = usedRange.Columns.Count;

                if (rowCount < 2)
                {
                    MessageBox.Show("Excel file does not contain enough rows.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Read header row (first row)
                var headers = new List<string>();
                for (int col = 1; col <= colCount; col++)
                {
                    var headerValue = (usedRange.Cells[1, col] as Range)?.Value2?.ToString() ?? string.Empty;
                    headers.Add(headerValue);
                }

                if (maxRows == 0) maxRows = rowCount - 1; // Process all rows if Max_Rows is 0
                else if (maxRows > rowCount - 1) maxRows = rowCount - 1; // Limit to available rows

                // Data rows - skip header row
                int batchSize = ConfigManager.Batch_Size;
                var batch = new List<Dictionary<string, string>>(batchSize);
                int processed = 0;
                // Process rows asynchronously
                await Task.Run(async () =>
                {
                    for (int row = 2; row <= rowCount; row++)
                    {
                        var rowData = new Dictionary<string, string>();
                        for (int col = 1; col <= colCount; col++)
                        {
                            var cellValue = (usedRange.Cells[row, col] as Range)?.Value2?.ToString() ?? string.Empty;
                            rowData[headers[col - 1]] = cellValue;
                        }
                        batch.Add(rowData);
                        processed++;

                        if (batch.Count == batchSize)
                        {
                            await ProcessBatchAsync(batch, excelApp);
                            batch.Clear();
                            GC.Collect();
                        }

                        if ((row - 1) >= maxRows)
                            break;
                    }

                    // Process any remaining rows
                    if (batch.Count > 0)
                    {
                        await ProcessBatchAsync(batch, excelApp);
                        batch.Clear();
                        GC.Collect();
                    }
                });
                Invoke(new System.Action(() =>
                {
                    lblActionDisplay.Text = "Generation completed.";
                }));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading Excel file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Release COM objects in reverse order of creation
                if (usedRange != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(usedRange);
                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (masterFile != null)
                {
                    masterFile.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(masterFile);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Invoke(new System.Action(() =>
                {
                    Cursor = Cursors.Default;
                    progressBar.Style = ProgressBarStyle.Blocks;
                    progressBar.Visible = false;
                    btnViewOutput.Visible = true;
                    btnViewOutput.Text = "View Output Folder";

                }));
            }
        }

        private async Task ProcessBatchAsync(List<Dictionary<string, string>> batch, Application excelApp)
        {
            foreach (var row in batch)
            {
                await ProcessRow(row, excelApp);
            }
        }

        private async Task ProcessRow(Dictionary<string, string> row, Application excelApp)
        {
            if (!row.TryGetValue(ColumnNames.SubjectCode, out var subjectCode) ||
                !row.TryGetValue(ColumnNames.StudyPeriod, out var studyPeriod))
            {
                MessageBox.Show($"Missing {ColumnNames.SubjectCode} or {ColumnNames.StudyPeriod}.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var sType = row.TryGetValue(ColumnNames.SubjectType, out var subjectType) ? subjectType : string.Empty;
            if (sType.ToLower().Trim() != ConfigManager.Coursework_Text.ToLower().Trim())
            {
                return; // Skip non-coursework subjects
            }

            Workbook workbook = null;
            Worksheet dataSheet = null;

            try
            {
                // Build safe filename
                var safeSubjectCode = string.Concat(subjectCode.Split(Path.GetInvalidFileNameChars())).ToLowerInvariant();
                var safeStudyPeriod = string.Concat(studyPeriod.Split(Path.GetInvalidFileNameChars())).ToLowerInvariant();
                var fileName = $"{safeSubjectCode}_{safeStudyPeriod}.xlsx";

                var outputDir = ConfigManager.Output_Location;
                Directory.CreateDirectory(outputDir);
                var targetPath = Path.Combine(outputDir, fileName);

                // Check if the file already exists
                if (File.Exists(targetPath))
                {
                    // Open the existing file
                    workbook = excelApp.Workbooks.Open(targetPath);
                }
                else
                {
                    // Copy template to writable location
                    File.Copy(ConfigManager.Template_File, targetPath, overwrite: true);

                    // Open the copied template file
                    workbook = excelApp.Workbooks.Open(targetPath);
                }

                // Try to get "Data" sheet, or create if missing
                dataSheet = null;
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    if (ws.Name == "Data")
                    {
                        dataSheet = ws;
                        break;
                    }
                }
                if (dataSheet == null)
                {
                    dataSheet = workbook.Worksheets.Add();
                    dataSheet.Name = "Data";
                }

                // Clear existing data
                dataSheet.Cells.Clear();

                // Write headers
                int col = 1;
                foreach (var header in row.Keys)
                {
                    dataSheet.Cells[1, col] = header;
                    col++;
                }

                // Write row values
                col = 1;
                foreach (var value in row.Values)
                {
                    dataSheet.Cells[2, col] = value;
                    col++;
                }

                // Save changes
                workbook.Save();

                // Extract total rows from lblMasterFileRowCount
                var totalRowsText = lblMasterFileRowCount.Text.Replace("Rows in Master File: ", "").Trim();
                if (!int.TryParse(totalRowsText, out var totalRows))
                {
                    totalRows = 0; // Fallback if parsing fails
                }

                // Update the label and listbox for each successfully processed file
                Invoke(new System.Action(() =>
                {
                    lblReport.Visible = true;
                    lblReport.Text = $"Generated {lstReport.Items.Count + 1} out of {totalRows - 1} files successfully...";
                    lstReport.Visible = true;
                    lstReport.Items.Add($"File {lstReport.Items.Count + 1}: {fileName}");
                }));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to create/update file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Release COM objects
                if (dataSheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(dataSheet);
                if (workbook != null)
                {
                    workbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                }
                // Do not quit or release excelApp here, as it is managed by the parent method
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void UnifyBtn_Click(object sender, EventArgs e)
        {
            var unifyFolder = ConfigManager.Unify_Folder;
            var doneFolder = ConfigManager.Done_Folder_Format;
            var reportPath = ConfigManager.Report_File_Format;

            // Replace datetime format (yyyyMMddHHmm)
            var timestamp = DateTime.Now.ToString("yyyyMMddHHmm");
            var reportFileName = ConfigManager.Report_File_Format.Replace("yyyyMMddHHmm", timestamp);
            reportPath = Path.Combine(unifyFolder, reportFileName);
            doneFolder = Path.Combine(unifyFolder, doneFolder.Replace("yyyyMMddHHmm", timestamp));

            //Create folder to store successfully processed files
            Directory.CreateDirectory(doneFolder);

            // Kill all running Excel processes before starting
            foreach (var process in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
            {
                try { process.Kill(); }
                catch { /* ignore if cannot kill */ }
            }

            Application excelApp = new Application();
            Workbook reportWb = excelApp.Workbooks.Add();
            Worksheet reportWs = (Worksheet)reportWb.Worksheets[1];

            int reportRow = 1;

            try
            {
                foreach (var file in Directory.GetFiles(unifyFolder, "*.xlsx", SearchOption.TopDirectoryOnly))
                {
                    Workbook srcWb = excelApp.Workbooks.Open(file);
                    Worksheet srcWs = (Worksheet)srcWb.Worksheets[ConfigManager.Workload_Main_Sheet];

                    // Mapping
                    reportWs.Cells[reportRow, 1] = (srcWs.Cells[3, 3] as Range)?.Value2?.ToString() ?? ""; // C3 -> A1
                    reportWs.Cells[reportRow, 2] = (srcWs.Cells[5, 3] as Range)?.Value2?.ToString() ?? ""; // C5 -> B1
                    reportWs.Cells[reportRow, 3] = (srcWs.Cells[3, 5] as Range)?.Value2?.ToString() ?? ""; // E3 -> C1

                    // Find START and END in column A
                    int startRow = 0, endRow = 0;
                    int lastRow = srcWs.UsedRange.Rows.Count;
                    for (int r = 1; r <= lastRow; r++)
                    {
                        var val = (srcWs.Cells[r, 1] as Range)?.Value2?.ToString();
                        if (val == "START") startRow = r + 1;
                        if (val == "END") { endRow = r - 1; break; }
                    }

                    // For loop row from START to END (column A)
                    for (int r = startRow; r <= endRow; r++)
                    {
                        var bVal = (srcWs.Cells[r, 2] as Range)?.Value2?.ToString();

                        if (!string.IsNullOrWhiteSpace(bVal))
                        {
                            // Repeat the header mappings for each copied row
                            reportWs.Cells[reportRow, 1] = (srcWs.Cells[3, 3] as Range)?.Value2?.ToString() ?? ""; // C3 -> A
                            reportWs.Cells[reportRow, 2] = (srcWs.Cells[5, 3] as Range)?.Value2?.ToString() ?? ""; // C5 -> B
                            reportWs.Cells[reportRow, 3] = (srcWs.Cells[3, 5] as Range)?.Value2?.ToString() ?? ""; // E3 -> C

                            // Copy column B value
                            reportWs.Cells[reportRow, 4] = bVal; // B(row) -> D
                            reportRow++;
                        }
                    }

                    srcWb.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(srcWs);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(srcWb);

                    // Move file to Done folder
                    var destFile = Path.Combine(doneFolder, Path.GetFileName(file));
                    File.Move(file, destFile);
                }

                reportWb.SaveAs(reportPath);
                MessageBox.Show($"Report generated: {reportPath}", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                reportWb.Close(false);
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(reportWs);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(reportWb);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void btnViewMaster_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe", Path.GetDirectoryName(txtMasterFile.Text));
        }

        private void btnViewTemplate_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe", Path.GetDirectoryName(txtTemplateFile.Text));
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            var masterFilePath = ConfigManager.Master_File;
            if (!File.Exists(masterFilePath))
            {
                MessageBox.Show("Master file not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Check if the template file exists
            var templateFilePath = ConfigManager.Template_File;
            if (!File.Exists(templateFilePath))
            {
                MessageBox.Show("Template file not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Application excelApp = null;
            Workbook masterFile = null;
            Worksheet worksheet = null;
            Range usedRange = null;

            try
            {
                excelApp = new Application();
                masterFile = excelApp.Workbooks.Open(masterFilePath);
                worksheet = (Worksheet)masterFile.Worksheets[1];
                usedRange = worksheet.UsedRange;

                // Remove all filters in the master file
                if (worksheet.AutoFilterMode)
                {
                    worksheet.AutoFilterMode = false;
                }

                int rowCount = usedRange.Rows.Count - 1;
                lblMasterFileRowCount.Text = $"Rows in Master File: {rowCount}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading master file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (usedRange != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(usedRange);
                if (worksheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                if (masterFile != null)
                {
                    masterFile.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(masterFile);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void btnCloseExcels_Click(object sender, EventArgs e)
        {
            // Get all running Excel processes
            var excelProcesses = Process.GetProcessesByName("EXCEL");

            // Attempt to close each process
            foreach (var process in excelProcesses)
            {
                try
                {
                    process.Kill();
                    process.WaitForExit(); // Ensure the process is terminated
                }
                catch (Exception ex)
                {
                    // Log or handle any errors while killing the process
                    MessageBox.Show($"Failed to close an Excel process: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            // Display a message indicating the operation is complete
            MessageBox.Show("All running Excel processes have been closed.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }

        private void btnViewOutput_Click(object sender, EventArgs e)
        {
            //Open the output folder in File Explorer
            Process.Start("explorer.exe", ConfigManager.Output_Location);
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            // Display a warning message to the user
            var result = MessageBox.Show(
                "Are you sure you want to close the application? Any ongoing generation or unification process will be terminated.",
                "Warning",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning
            );

            // If the user selects "No", cancel the close operation
            if (result == DialogResult.No)
            {
                return;
            }

            // Check if the current process is generating or unifying
            if (lblActionDisplay.Visible && lblActionDisplay.Text.Contains("Generating") || lblActionDisplay.Text.Contains("Unifying"))
            {
                // Close all running Excel processes
                var excelProcesses = Process.GetProcessesByName("EXCEL");
                foreach (var process in excelProcesses)
                {
                    try
                    {
                        process.Kill();
                        process.WaitForExit(); // Ensure the process is terminated
                    }
                    catch (Exception ex)
                    {
                        // Log or handle any errors while killing the process
                        MessageBox.Show($"Failed to close an Excel process: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }

            // Close the form
            this.Close();
        }
    }
}
