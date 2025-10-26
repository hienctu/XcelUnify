using Microsoft.Office.Interop.Excel;
using XcelUnify.Helpers;
using Range = Microsoft.Office.Interop.Excel.Range;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace XcelUnify
{
    public partial class Main : Form
    {
        private string rptFolderPath;
        private string tempStaffUpdateFolder;

        public Main()
        {
            InitializeComponent();
            txtMasterFile.Text = ConfigManager.Master_File;

            txtTemplateFile.Text = ConfigManager.GetTemplateFile(ConfigManager.Coursework_Text);
            txtTemplateFile.ReadOnly = true;

            txtResearchTemplateFile.Text = ConfigManager.GetTemplateFile(ConfigManager.Research_Text);
            txtResearchTemplateFile.ReadOnly = true;

            txtDualCampusTemplateFile.Text = ConfigManager.GetTemplateFile(ConfigManager.DualCampus_Text);
            txtDualCampusTemplateFile.ReadOnly = true;

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
            toolTips.SetToolTip(btnGenerate, "Generate individual workload files for only Coursework subject based on the master file and template SAFES file");
            toolTips.SetToolTip(UnifyBtn, "Unify all individual workload files in the specified folder into a single report file and move processed files to a 'Done' folder");

        }

        private async void btnGenerate_Click(object sender, EventArgs e)
        {
            // Change the cursor to "Wait"
            lstReport.Items.Clear();
            Cursor = Cursors.WaitCursor;
            int fromRow = ConfigManager.Generate_From_Row;
            int toRow = ConfigManager.Generate_To_Row;
            int maxRows = toRow - fromRow + 1;

            Invoke(new System.Action(() =>
            {
                lblActionDisplay.Visible = true;
                lblActionDisplay.Text = String.Format("Generating workload files...(from row {0} to row {1} in master data file)", fromRow, toRow);
                progressBar.Visible = true;
                progressBar.Style = ProgressBarStyle.Marquee;
            }));



            int rowCount = 0;
            int colCount = 0;

            // 1. Create temp working folder
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string tempWorkFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Working", $"TempWork_{timestamp}");
            Directory.CreateDirectory(tempWorkFolder);

            var masterFilePath = ConfigManager.Master_File;
            string tempMasterFile = Path.Combine(tempWorkFolder, Path.GetFileName(masterFilePath));
            File.Copy(masterFilePath, tempMasterFile, true);

            // Copy all template files to the temp folder
            string[] templateFiles = { "standard-template.xlsx", "research-template.xlsx", "dual-template.xlsx" };
            foreach (var templateFile in templateFiles)
            {
                string source = Path.Combine(ConfigManager.Template_File_Path, templateFile);
                string dest = Path.Combine(tempWorkFolder, templateFile);
                if (File.Exists(source))
                {
                    File.Copy(source, dest, true);
                }
            }


            string outputDir = ConfigManager.Output_Location;
            string tempOutputDir = Path.Combine(tempWorkFolder, "Output");
            Directory.CreateDirectory(tempOutputDir);
            foreach (var file in Directory.GetFiles(outputDir, "*.xlsx", SearchOption.TopDirectoryOnly))
            {
                string dest = Path.Combine(tempOutputDir, Path.GetFileName(file));
                File.Copy(file, dest, true);
            }

            var templateFilePath = string.Empty;
            var masterHeaderRow = ConfigManager.Master_First_Data_Row - 1;

            Application excelApp = null;
            Workbook masterFile = null;
            Worksheet worksheet = null;
            Range usedRange = null;

            try
            {
                excelApp = new Application();
                masterFile = excelApp.Workbooks.Open(tempMasterFile);
                worksheet = (Worksheet?)masterFile.Worksheets[1];
                usedRange = worksheet.UsedRange;
                rowCount = usedRange.Rows.Count - masterHeaderRow;
                colCount = usedRange.Columns.Count;

                if (rowCount < ConfigManager.Master_First_Data_Row)
                {
                    MessageBox.Show("Excel file does not contain enough rows.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Read header row (first row)
                var headers = new List<string>();
                for (int col = 1; col <= colCount; col++)
                {
                    var headerValue = (usedRange.Cells[masterHeaderRow, col] as Range)?.Value2?.ToString() ?? string.Empty;
                    headers.Add(headerValue);
                }

                // Data rows - skip header row
                int batchSize = ConfigManager.Batch_Size;
                var batch = new List<Dictionary<string, string>>(batchSize);
                int processed = 0;
                // Process rows asynchronously
                await Task.Run(async () =>
                {
                    for (int row = fromRow; row <= toRow; row++)
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
                            await ProcessBatchAsync(batch, excelApp, maxRows, tempWorkFolder, tempOutputDir);
                            batch.Clear();
                            GC.Collect();
                        }
                    }

                    // Process any remaining rows
                    if (batch.Count > 0)
                    {
                        await ProcessBatchAsync(batch, excelApp, maxRows, tempWorkFolder, tempOutputDir);
                        batch.Clear();
                        GC.Collect();
                    }
                });

                //Move all files from tempOutputDir to outputDir
                if (Directory.Exists(tempOutputDir))
                {
                    foreach (var file in Directory.GetFiles(tempOutputDir, "*.xlsx", SearchOption.TopDirectoryOnly))
                    {
                        var destFile = Path.Combine(outputDir, Path.GetFileName(file));
                        File.Copy(file, destFile, true);
                    }
                }




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

            //Now delete temp working folder
            try
            {
                if (Directory.Exists(tempWorkFolder))
                {
                    Directory.Delete(tempWorkFolder, true); // true = recursive delete
                }
            }
            catch (Exception ex)
            {
                // Optionally log or show a warning, but do not block the user
                Debug.WriteLine($"Failed to delete temp working folder: {ex.Message}");
            }
        }

        private async Task ProcessBatchAsync(List<Dictionary<string, string>> batch, Application excelApp, int numberRowsToGenerate, string tempWorkFolder, string tempOutputDir)
        {
            foreach (var row in batch)
            {
                await ProcessRow(row, excelApp, numberRowsToGenerate, tempWorkFolder, tempOutputDir);
            }
        }

        private async Task ProcessRow(Dictionary<string, string> row, Application excelApp, int numberRowsToGenerate, string tempWorkFolder, string tempOutputDir)
        {
            if (!row.TryGetValue(ColumnNames.SubjectCode, out var subjectCode) ||
                !row.TryGetValue(ColumnNames.StudyPeriod, out var studyPeriod))
            {
                MessageBox.Show($"Missing {ColumnNames.SubjectCode} or {ColumnNames.StudyPeriod}.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var sType = row.TryGetValue(ColumnNames.Category, out var category) ? category : string.Empty;
            var templateFile = ConfigManager.GetTemplateFile(sType, tempWorkFolder);
            if (string.IsNullOrEmpty(templateFile))
            {
                return; // Skip - cannot find template file
            }

            Workbook workbook = null;
            Worksheet dataSheet = null;

            var safeSubjectCode = string.Concat(subjectCode.Split(Path.GetInvalidFileNameChars())).ToLowerInvariant();
            var safeStudyPeriod = string.Concat(studyPeriod.Split(Path.GetInvalidFileNameChars())).ToLowerInvariant();
            var fileName = $"{safeSubjectCode}_{safeStudyPeriod}.xlsx";

            try
            {
                var targetPath = Path.Combine(tempOutputDir, fileName);

                // Check if the file already exists
                if (File.Exists(targetPath))
                {
                    // Open the existing file
                    workbook = excelApp.Workbooks.Open(targetPath);
                }
                else
                {
                    // Copy template to writable location
                    File.Copy(templateFile, targetPath, overwrite: true);

                    // Open the copied template file
                    workbook = excelApp.Workbooks.Open(targetPath);
                }

                //Start unlock the file
                workbook.Unprotect(ConfigManager.Template_File_Password);

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

                dataSheet.Visible = XlSheetVisibility.xlSheetVeryHidden;

                //Protect the workbook again
                workbook.Protect(ConfigManager.Template_File_Password);

                // Save changes
                workbook.Save();

                var totalRows = numberRowsToGenerate > 0 ? numberRowsToGenerate : 1;

                // Update the label and listbox for each successfully processed file
                Invoke(new System.Action(() =>
                {
                    lblReport.Visible = true;
                    lblReport.Text = $"Generated {lstReport.Items.Count + 1} out of {totalRows} files successfully...";
                    lstReport.Visible = true;
                    lstReport.Items.Add($"File {lstReport.Items.Count + 1}: {fileName}");
                }));
            }
            catch (COMException comEx) when (comEx.Message.Contains("password"))
            {
                // Specific error for incorrect password
                Invoke(new System.Action(() =>
                {
                    lstReport.Items.Add($"File {fileName} could not be unlocked with the provided password. Skipping...");
                }));
            }
            catch (Exception ex)
            {
                // Generic error handling
                Invoke(new System.Action(() =>
                {
                    lstReport.Items.Add($"Error - File {fileName} encountered an error. Skipping...");
                }));
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

        private async void UnifyBtn_Click(object sender, EventArgs e)
        {
            lstReport.Items.Clear();
            Cursor = Cursors.WaitCursor;

            Invoke(new System.Action(() =>
            {
                lblActionDisplay.Visible = true;
                lblActionDisplay.Text = "Unifying SAFES workload files...";
                progressBar.Visible = true;
                progressBar.Style = ProgressBarStyle.Marquee;
                lstReport.Visible = true;
            }));

            var unifyFolder = ConfigManager.Unify_Folder;
            // 1. Create temp working folder and copy all files from unifyFolder to tempWorkFolder
            string timestampHHMMSS = DateTime.Now.ToString("yyyyMMddHHmmss");
            string tempWorkFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Working", $"Unify_TempWork_{timestampHHMMSS}");
            Directory.CreateDirectory(tempWorkFolder);

            //copy all files from unifyFolder to tempWorkFolder
            foreach (var file in Directory.GetFiles(unifyFolder, "*.xlsx", SearchOption.TopDirectoryOnly))
            {
                var destFile = Path.Combine(tempWorkFolder, Path.GetFileName(file));
                File.Copy(file, destFile, true);
            }


            var doneFolder = ConfigManager.Done_Folder_Format;
            var reportPath = ConfigManager.Report_File_Format;

            // Replace datetime format (yyyyMMddHHmm)
            var timestamp = DateTime.Now.ToString("yyyyMMddHHmm");
            var reportFileName = ConfigManager.Report_File_Format.Replace("yyyyMMddHHmm", timestamp);
            Directory.CreateDirectory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "UnifyRpt"));
            doneFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "UnifyRpt", doneFolder.Replace("yyyyMMddHHmm", timestamp));
            reportPath = Path.Combine(doneFolder, reportFileName);
            rptFolderPath = doneFolder;
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
                // Add headers to the report
                string[] headers = new string[]
                {
                    "Subject Code", "Subject Title", "Study Period", "Est. Enrolment", "% Allocation",
                    "Staff Name", "Coordinator", "Lecture Initial", "Lecture Repeat", "Tute/WS Initial", "Tute/WS Repeat",
                    "Practical Initial", "Practical Repeat", "FieldTrip/Excursion", "Marking"
                };

                for (int col = 1; col <= headers.Length; col++)
                {
                    reportWs.Cells[reportRow, col] = headers[col - 1];
                }

                // Apply formatting: gray background and bold text
                Range headerRange = reportWs.Range[reportWs.Cells[reportRow, 1], reportWs.Cells[reportRow, headers.Length]];
                headerRange.Interior.Color = ColorTranslator.ToOle(Color.LightGray);
                headerRange.Font.Bold = true;

                reportRow++; // Move to the next row for data
                await Task.Run(async () =>
                {
                    var filesCount = Directory.GetFiles(tempWorkFolder, "*.xlsx", SearchOption.TopDirectoryOnly).Count();

                    foreach (var file in Directory.GetFiles(tempWorkFolder, "*.xlsx", SearchOption.TopDirectoryOnly))
                    {
                        try
                        {
                            Workbook srcWb = excelApp.Workbooks.Open(file);
                            Worksheet srcWs = (Worksheet)srcWb.Worksheets[ConfigManager.Workload_Main_Sheet];
                            // Find START and END in column A
                            int startRow = 0,
                                endRow = 0;

                            int otherStaffStartRow = 0;
                            int otherStaffEndRow = 0;

                            decimal allocatedPercent = 0;

                            var allocatedValue = (srcWs.Range[ConfigManager.Allocated_Overall_Address] as Range)?.Value2?.ToString();
                            decimal resultAllocation;
                            allocatedPercent = decimal.TryParse(allocatedValue, out resultAllocation) ? Math.Round(resultAllocation, 1) : 0;
                            // Assuming labels are in column A
                            for (int row = 1; row <= srcWs.UsedRange.Rows.Count; row++)
                            {
                                var cellValue = (srcWs.Cells[row, 2] as Range)?.Value2?.ToString();

                                if (cellValue != null)
                                {
                                    if (cellValue.Trim().ToLower() == ConfigManager.SafesStaff_Label.Trim().ToLower())
                                    {
                                        startRow = row + 2; // Start row is the row after the label
                                    }
                                    else if (cellValue.Contains(ConfigManager.TotalHrs_Label, StringComparison.OrdinalIgnoreCase)
                                                && startRow > 0 && row > startRow)
                                    {
                                        endRow = row - 1; // End row is the row before this label
                                        break;
                                    }
                                }
                            }
                            for (int row = endRow + 1; row <= srcWs.UsedRange.Rows.Count; row++)
                            {
                                var cellValue = (srcWs.Cells[row, 2] as Range)?.Value2?.ToString();

                                if (cellValue != null)
                                {
                                    if (cellValue.Trim().ToLower() == ConfigManager.OtherStaff_Label.Trim().ToLower())
                                    {
                                        otherStaffStartRow = row + 2; // Start row is the row after the label
                                    }
                                    else if (cellValue.Contains(ConfigManager.TotalHrs_Label, StringComparison.OrdinalIgnoreCase)
                                                && otherStaffStartRow > 0
                                                && otherStaffStartRow > startRow
                                                && row > otherStaffStartRow)
                                    {
                                        otherStaffEndRow = row - 1; // End row is the row before this label
                                        break; // Exit loop once both labels are found
                                    }
                                }
                            }

                            if (startRow == 0 || endRow == 0)
                            {
                                throw new Exception("Could not find SafesStaffLabel or TotalHrsLabel in the worksheet.");
                            }


                            // For loop row from START to END (column A)
                            for (int r = startRow; r <= endRow; r++)
                            {
                                var bVal = (srcWs.Cells[r, 2] as Range)?.Value2?.ToString();

                                if (!string.IsNullOrWhiteSpace(bVal))
                                {
                                    // Repeat the header mappings for each copied row
                                    reportWs.Cells[reportRow, 1] = (srcWs.Cells[3, 3] as Range)?.Value2?.ToString() ?? ""; // C3 - Code
                                    reportWs.Cells[reportRow, 2] = (srcWs.Cells[3, 4] as Range)?.Value2?.ToString() ?? ""; // C5 - Name
                                    reportWs.Cells[reportRow, 3] = (srcWs.Cells[7, 3] as Range)?.Value2?.ToString() ?? ""; // C7 - Timing
                                    reportWs.Cells[reportRow, 4] = (srcWs.Cells[8, 3] as Range)?.Value2?.ToString() ?? ""; // C8 - Enrolment
                                    reportWs.Cells[reportRow, 5] = allocatedPercent; // % Allocation

                                    if (allocatedPercent < 100)
                                    {
                                        for (int col = 1; col <= 5; col++)
                                        {
                                            var cell = reportWs.Cells[reportRow, col] as Range;
                                            if (cell != null)
                                            {
                                                cell.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightYellow);
                                            }
                                        }
                                    }

                                    int lastUsedColumn = srcWs.UsedRange.Columns.Count; // Get the last used column in the source worksheet
                                    reportWs.Cells[reportRow, 6] = (srcWs.Cells[r, 2] as Range)?.Value2?.ToString() ?? ""; // Staff name

                                    for (int workloadFileCol = 3; workloadFileCol <= lastUsedColumn; workloadFileCol++)
                                    {
                                        reportWs.Cells[reportRow, workloadFileCol + 4] = (srcWs.Cells[r, workloadFileCol] as Range)?.Value2?.ToString() ?? ""; // Adjust column index for the report worksheet
                                    }
                                    reportRow++;
                                }
                            }

                            for (int r = otherStaffStartRow; r <= otherStaffEndRow; r++)
                            {
                                var bVal = (srcWs.Cells[r, 2] as Range)?.Value2?.ToString();

                                if (!string.IsNullOrWhiteSpace(bVal))
                                {
                                    int lastUsedColumn = srcWs.UsedRange.Columns.Count; // Get the last used column in the source worksheet
                                    bool hasValue = false;
                                    for (int workloadFileCol = 3; workloadFileCol <= lastUsedColumn; workloadFileCol++)
                                    {
                                        var cellValue = (srcWs.Cells[r, workloadFileCol] as Range)?.Value2?.ToString();
                                        if (!string.IsNullOrWhiteSpace(cellValue))
                                        {
                                            hasValue = true;
                                            break;
                                        }
                                    }
                                    if (hasValue)
                                    {
                                        // Repeat the header mappings for each copied row
                                        reportWs.Cells[reportRow, 1] = (srcWs.Cells[3, 3] as Range)?.Value2?.ToString() ?? ""; // C3 - Code
                                        reportWs.Cells[reportRow, 2] = (srcWs.Cells[3, 4] as Range)?.Value2?.ToString() ?? ""; // C5 - Name
                                        reportWs.Cells[reportRow, 3] = (srcWs.Cells[7, 3] as Range)?.Value2?.ToString() ?? ""; // C7 - Timing
                                        reportWs.Cells[reportRow, 4] = (srcWs.Cells[8, 3] as Range)?.Value2?.ToString() ?? ""; // C8 - Enrolment
                                        reportWs.Cells[reportRow, 5] = allocatedPercent; // % Allocation

                                        if (allocatedPercent < 100)
                                        {
                                            for (int col = 1; col <= 5; col++)
                                            {
                                                var cell = reportWs.Cells[reportRow, col] as Range;
                                                if (cell != null)
                                                {
                                                    cell.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightYellow);
                                                }
                                            }
                                        }

                                        reportWs.Cells[reportRow, 6] = (srcWs.Cells[r, 2] as Range)?.Value2?.ToString() ?? ""; // Staff name

                                        for (int workloadFileCol = 3; workloadFileCol <= lastUsedColumn; workloadFileCol++)
                                        {
                                            reportWs.Cells[reportRow, workloadFileCol + 4] = (srcWs.Cells[r, workloadFileCol] as Range)?.Value2?.ToString() ?? "";
                                        }
                                        reportRow++;
                                    }
                                }
                            }

                            srcWb.Close(false);
                            Marshal.ReleaseComObject(srcWs);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(srcWb);

                            // Move file to Done folder
                            var destFile = Path.Combine(doneFolder, Path.GetFileName(file));
                            File.Move(file, destFile);
                            // Update the label and listbox for each successfully processed file
                            Invoke(new System.Action(() =>
                            {
                                lblReport.Visible = true;
                                lblReport.Text = $"Unified {lstReport.Items.Count + 1} out of {filesCount} files successfully...";
                                lstReport.Visible = true;
                                lstReport.Items.Add($"Collected data from file {lstReport.Items.Count + 1}: {Path.GetFileName(file)}");
                            }));
                        }
                        catch (Exception ex)
                        {
                            // Log error and continue with next file
                            Invoke(new System.Action(() =>
                            {
                                lstReport.Items.Add($"Error - File {Path.GetFileName(file)} encountered an error. Skipping...");
                            }));
                        }
                    }
                }); //End of Task.Run


                reportWs.Columns.AutoFit();
                reportWb.SaveAs(reportPath);
                Invoke(new System.Action(() =>
                {
                    lblActionDisplay.Text = "Unification process completed successfully.";
                }));

                //Now delete temp working folder
                try
                {
                    if (Directory.Exists(tempWorkFolder))
                    {
                        Directory.Delete(tempWorkFolder, true); // true = recursive delete
                    }
                }
                catch (Exception ex)
                {
                    // Optionally log or show a warning, but do not block the user
                    Debug.WriteLine($"Failed to delete temp working folder: {ex.Message}");
                }
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

                Invoke(new System.Action(() =>
                {
                    Cursor = Cursors.Default;
                    progressBar.Style = ProgressBarStyle.Blocks;
                    progressBar.Visible = false;
                    btnViewOutput.Visible = true;
                    btnViewOutput.Text = "View Report Folder";

                }));
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
                MessageBox.Show("Master Data file not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Check if the template file exists
            var templateFilePath = ConfigManager.GetTemplateFile(ConfigManager.Coursework_Text);
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

                int rowCount = usedRange.Rows.Count - ConfigManager.Master_First_Data_Row + 1;
                lblMasterFileRowCount.Text = $"Rows in Master File: {rowCount}";
                if (ConfigManager.Generate_To_Row == 0)
                {
                    ConfigManager.Generate_To_Row = usedRange.Rows.Count;
                }
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
            string folderPath;

            if (btnViewOutput.Text.Contains("Staff Update") && !string.IsNullOrEmpty(tempStaffUpdateFolder))
            {
                folderPath = tempStaffUpdateFolder;
            }
            else if (btnViewOutput.Text.Contains("Output"))
            {
                folderPath = ConfigManager.Output_Location;
            }
            else
            {
                folderPath = rptFolderPath;
            }

            if (!string.IsNullOrEmpty(folderPath) && Directory.Exists(folderPath))
            {
                Process.Start("explorer.exe", folderPath);
            }
            else
            {
                MessageBox.Show("The specified folder does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

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

        private void button1_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe", Path.GetDirectoryName(txtResearchTemplateFile.Text));
        }

        private void btnViewDualCampusTemplate_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe", Path.GetDirectoryName(txtDualCampusTemplateFile.Text));
        }

        private async void btnUpdateStaffList_Click(object sender, EventArgs e)
        {
            lstReport.Items.Clear();
            Cursor = Cursors.WaitCursor;

            Invoke(new System.Action(() =>
            {
                lblActionDisplay.Visible = true;
                lblActionDisplay.Text = "Preparing to update staff list in SAFES workload files and three templates...";
                progressBar.Visible = true;
                progressBar.Style = ProgressBarStyle.Marquee;
                lstReport.Visible = true;
            }));


            /* Open after testing */
            
            // 1. Create temp working folder
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string tempWorkFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Working", $"UpdateStaff_TempWork_{timestamp}");
            Directory.CreateDirectory(tempWorkFolder);
            Directory.CreateDirectory(Path.Combine(tempWorkFolder, "Data"));

            // 2. Copy master data file to temp folder
            var masterFilePath = ConfigManager.Master_File;
            string tempMasterFile = Path.Combine(tempWorkFolder, "Data", Path.GetFileName(masterFilePath));
            File.Copy(masterFilePath, tempMasterFile, true);

            // 3. Copy all template files to the temp folder
            string[] templateFiles = { "standard-template.xlsx", "research-template.xlsx", "dual-template.xlsx" };
            foreach (var templateFile in templateFiles)
            {
                string source = Path.Combine(ConfigManager.Template_File_Path, templateFile);
                string dest = Path.Combine(tempWorkFolder, "Data", templateFile);
                if (File.Exists(source))
                {
                    File.Copy(source, dest, true);
                }
            }

            // 4. Copy all generated files from output folder to temp folder
            string outputDir = ConfigManager.Output_Location;
            foreach (var file in Directory.GetFiles(outputDir, "*.xlsx", SearchOption.TopDirectoryOnly))
            {
                string dest = Path.Combine(tempWorkFolder, Path.GetFileName(file));
                File.Copy(file, dest, true);
            }

            /* Testing only
            string tempWorkFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Working", $"UpdateStaff_TempWork_20251026120226");
            tempStaffUpdateFolder = tempWorkFolder;
            var masterFilePath = ConfigManager.Master_File;
            string tempMasterFile = Path.Combine(tempWorkFolder, "Data", Path.GetFileName(masterFilePath));
            */

            // Kill all running Excel processes before starting
            foreach (var process in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
            {
                try { process.Kill(); }
                catch { /* ignore if cannot kill */ }
            }

            // 4. (Optional) Add your staff list update logic here, working in tempWorkFolder
            int noStaffInMaster = 0;
            List<string> staffNames = new List<string>();

            Application excelApp = null;
            Workbook masterWb = null;
            Worksheet staffSheet = null;

            try
            {
                await Task.Run(async () =>
                {
                    var filesCount = Directory.GetFiles(tempWorkFolder, "*.xlsx", SearchOption.TopDirectoryOnly).Count();

                    excelApp = new Application();
                    masterWb = excelApp.Workbooks.Open(tempMasterFile);
                    staffSheet = masterWb.Worksheets["Staff List"] as Worksheet;
                    if (staffSheet == null)
                    {
                        MessageBox.Show("Sheet 'Staff List' not found in master file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    Range usedRangeMaster = staffSheet.UsedRange;
                    int lastRow = usedRangeMaster.Rows.Count;

                    // Start from row 2 (row 1 is header)
                    for (int row = 2; row <= lastRow; row++)
                    {
                        var value = (usedRangeMaster.Cells[row, 1] as Range)?.Value2?.ToString();
                        if (!string.IsNullOrWhiteSpace(value))
                        {
                            staffNames.Add(value);
                            noStaffInMaster++;
                        }
                    }

                    /* Update three templates first */
                    // Update staff list in all templates
                    // 1. Gather all files to update
                    string[] templateFilesToUpdate = { "standard-template.xlsx", "research-template.xlsx", "dual-template.xlsx" };
                    var filesToUpdate = new List<string>();

                    // Add templates from Data subfolder
                    foreach (var templateFileName in templateFilesToUpdate)
                    {
                        var templatePath = Path.Combine(tempWorkFolder, "Data", templateFileName);
                        if (File.Exists(templatePath))
                            filesToUpdate.Add(templatePath);
                    }

                    // Add all .xlsx files in tempWorkFolder (excluding templates if desired)
                    var allXlsxFiles = Directory.GetFiles(tempWorkFolder, "*.xlsx", SearchOption.TopDirectoryOnly);
                    foreach (var file in allXlsxFiles)
                    {
                        if (!filesToUpdate.Contains(file)) // Avoid double-processing templates
                            filesToUpdate.Add(file);
                    }
                    // 2. Process each file
                    //Testing - Take 10
                    //foreach (var templateFileName in filesToUpdate.Take(10))
                    foreach (var templateFileName in filesToUpdate)
                    {

                        Workbook templateWb = null;
                        Worksheet staffListSheet = null;
                        Worksheet srcWs = null;
                        Range usedRange = null;

                        try
                        {
                            excelApp = new Application();
                            templateWb = excelApp.Workbooks.Open(templateFileName);
                            templateWb.Unprotect(ConfigManager.Template_File_Password);

                            staffListSheet = templateWb.Worksheets[ConfigManager.StaffList_Sheet_Name] as Worksheet;
                            staffListSheet.Unprotect(ConfigManager.Template_File_Password);

                            if (staffListSheet == null)
                            {
                                MessageBox.Show($"Sheet '{ConfigManager.StaffList_Sheet_Name}' not found in {templateFileName}.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                continue;
                            }

                            // Count the existing staff entries from E4 downwards
                            usedRange = staffListSheet.UsedRange;
                            int existingStaffCount = 0;
                            for (int r = 4; r <= usedRange.Rows.Count; r++)
                            {
                                var cellValue = (usedRange.Cells[r, 5] as Range)?.Value2?.ToString();
                                if (!string.IsNullOrWhiteSpace(cellValue))
                                {
                                    existingStaffCount++;
                                }
                                else
                                {
                                    break; // Stop counting when an empty cell is found
                                }
                            }

                            if (existingStaffCount > noStaffInMaster)
                            {
                                var result = MessageBox.Show(
                                    $"The template '{templateFileName}' has {existingStaffCount} staff entries, which is more than the {noStaffInMaster} entries in the master file. Do you want to proceed with the update? Extra entries will be removed.",
                                    "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                if (result == DialogResult.No)
                                {
                                    return; // exit this method without making changes    
                                }
                            }

                            // Clear existing entries from E4 downwards
                            staffListSheet.Range["E4:E" + (existingStaffCount + 3)].ClearContents();

                            // Write new staff names starting from E4
                            for (int i = 0; i < staffNames.Count; i++)
                            {
                                staffListSheet.Cells[i + 4, 5] = staffNames[i]; // Column E is the 5th column
                            }

                            // Update the text in E2 - updated as at today dd/mm/yyyy
                            staffListSheet.Cells[2, 5] = $"Updated as at {DateTime.Now:dd/MM/yyyy}";

                            // Protect the workbook again
                            staffListSheet.Protect(ConfigManager.Template_File_Password);

                            // Hide the Staff List sheet again
                            staffListSheet.Visible = XlSheetVisibility.xlSheetHidden;

                            // Save and close the template
                            templateWb.Save();

                            //Start updating data validation
                            srcWs = (Worksheet)templateWb.Worksheets[ConfigManager.Workload_Main_Sheet];
                            srcWs.Unprotect(ConfigManager.Template_File_Password);
                            // Find START and END in column A
                            int startRow = 0,
                                endRow = 0;

                            // Assuming labels are in column A
                            for (int row = 1; row <= srcWs.UsedRange.Rows.Count; row++)
                            {
                                var cellValue = (srcWs.Cells[row, 2] as Range)?.Value2?.ToString();

                                if (cellValue != null)
                                {
                                    if (cellValue.Trim().ToLower() == ConfigManager.SafesStaff_Label.Trim().ToLower())
                                    {
                                        startRow = row + 2; // Start row is the row after the label
                                    }
                                    else if (cellValue.Contains(ConfigManager.TotalHrs_Label, StringComparison.OrdinalIgnoreCase)
                                                && startRow > 0 && row > startRow)
                                    {
                                        endRow = row - 1; // End row is the row before this label
                                        break;
                                    }
                                }
                            }

                            // Update data validation for staff names in column B from startRow to endRow
                            for (int r = startRow; r <= endRow; r++)
                            {
                                Range cell = srcWs.Cells[r, 2] as Range; // Column B
                                if (cell != null)
                                {
                                    // Create the data validation formula
                                    string formula = $"='{ConfigManager.StaffList_Sheet_Name}'!$E$4:$E${staffNames.Count + 3}";
                                    // Add data validation
                                    cell.Validation.Delete(); // Remove any existing validation
                                    cell.Validation.Add(
                                        XlDVType.xlValidateList,
                                        XlDVAlertStyle.xlValidAlertStop,
                                        XlFormatConditionOperator.xlBetween,
                                        formula,
                                        Type.Missing);
                                    cell.Validation.IgnoreBlank = true;
                                    cell.Validation.InCellDropdown = true;

                                }
                            }

                            srcWs.Protect(ConfigManager.Template_File_Password);
                            templateWb.Protect(ConfigManager.Template_File_Password);
                            templateWb.Save();

                            Invoke(new System.Action(() =>
                            {
                                lblReport.Visible = true;
                                lblReport.Text = $"Updated {lstReport.Items.Count + 1} out of {filesCount} files and 3 templates successfully...";
                                lstReport.Visible = true;
                                lstReport.Items.Add($"Updating {lstReport.Items.Count + 1}: {Path.GetFileName(templateFileName)}");
                            }));
                        }
                        catch (Exception ex)
                        {
                            Invoke(new System.Action(() =>
                            {
                                lstReport.Items.Add($"Error - File {Path.GetFileName(templateFileName)} encountered an error. Skipping...");
                            }));
                        }
                        finally
                        {
                            if (staffListSheet != null) Marshal.ReleaseComObject(staffListSheet);
                            if (templateWb != null)
                            {
                                templateWb.Close(false);
                                Marshal.ReleaseComObject(templateWb);
                            }
                        }
                    }

                    // Release COM objects - master file
                    if (usedRangeMaster != null) Marshal.ReleaseComObject(usedRangeMaster);

                }); // End of Task.Run
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error reading 'Staff List': {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (staffSheet != null) Marshal.ReleaseComObject(staffSheet);
                if (masterWb != null)
                {
                    masterWb.Close(false);
                    Marshal.ReleaseComObject(masterWb);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            // 5. Clean up UI
            Invoke(new System.Action(() =>
            {
                lblActionDisplay.Text = "Staff list update preparation completed.";
                Cursor = Cursors.Default;
                progressBar.Style = ProgressBarStyle.Blocks;
                progressBar.Visible = false;

                //need to view the button and when click on it open the temp folder
                btnViewOutput.Visible = true;
                btnViewOutput.Text = "View Staff Update Temp Folder";
            }));


            // 6. (Optional) Clean up temp folder if needed
            // try
            // {
            //     if (Directory.Exists(tempWorkFolder))
            //     {
            //         Directory.Delete(tempWorkFolder, true);
            //     }
            // }
            // catch (Exception ex)
            // {
            //     Debug.WriteLine($"Failed to delete temp working folder: {ex.Message}");
            // }

        }

        private void btnUploadStaffUpdate_Click(object sender, EventArgs e)
        {
            //Confirm user that are you sure to copy files from Staff Update Temp Folder overwrite existing files in SharePoint
            //In message box, we show 2 hyperlinks of the Staff Update Temp Folder and SharePoint Output Location
            var tempFolder = tempStaffUpdateFolder;
            // Ensure tempFolder is set - if tempStaffUpdateFolder is null or empty, show error
            if (string.IsNullOrEmpty(tempFolder))
            {
                MessageBox.Show("No Staff Update Temp Folder found. Only run this upload only after running Update Staff List.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var sharepointFolder = ConfigManager.Output_Location;

            using (var dlg = new HyperlinkForm(tempFolder, sharepointFolder))
            {
                var result = dlg.ShowDialog(this);
                if (result == DialogResult.Yes)
                {
                    // Proceed with the upload
                    //3 template files in Data subfolder need to be copied to TemplateFilePath
                    try
                    {
                        //rename the temp folder to indicate upload completed
                        string completedFolder = Path.Combine(Path.GetDirectoryName(tempFolder), tempFolder + "-Completed", Path.GetFileName(tempFolder));
                        Directory.CreateDirectory(Path.GetDirectoryName(completedFolder));
                        Directory.CreateDirectory(Path.Combine(completedFolder, "Data"));

                        // Also copy the 3 template files from Data subfolder
                        string dataSubfolder = Path.Combine(tempFolder, "Data");
                        string[] templateFiles = { "standard-template.xlsx", "research-template.xlsx", "dual-template.xlsx" };
                        foreach (var templateFile in templateFiles)
                        {
                            string source = Path.Combine(dataSubfolder, templateFile);
                            string dest = Path.Combine(ConfigManager.Template_File_Path, templateFile);
                            if (File.Exists(source))
                            {
                                File.Copy(source, dest, true);
                                //Move to complete data
                                string destCompleted = Path.Combine(completedFolder, "Data", templateFile);
                                File.Move(source, destCompleted, true);
                            }
                        }

                        //Copy all files from tempFolder to sharepointFolder
                        foreach (var file in Directory.GetFiles(tempFolder, "*.xlsx", SearchOption.TopDirectoryOnly))
                        {
                            var destFile = Path.Combine(sharepointFolder, Path.GetFileName(file));
                            File.Copy(file, destFile, true); // true = overwrite existing files
                            //Move to completed folder
                            var destFileCompleted = Path.Combine(completedFolder, Path.GetFileName(file));
                            File.Move(file, destFileCompleted, true);
                        }

                        MessageBox.Show("All files have been successfully uploaded to SharePoint Output Location.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error during upload: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                // If No, simply return
                else
                {
                    return;
                }
            }
        }
    }
}
