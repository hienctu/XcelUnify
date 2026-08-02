using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using XcelUnify.Helpers;
using Application = Microsoft.Office.Interop.Excel.Application;
using Range = Microsoft.Office.Interop.Excel.Range;

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
            txtMasterDashboard.Text = ConfigManager.Master_Dashboard_File;

            txtUnifiedMasterFile.Text = ConfigManager.Unified_Master_File;
            
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
            var unifiedMasterFilePath = ConfigManager.Unified_Master_File;
            string tempMasterFile = Path.Combine(tempWorkFolder, Path.GetFileName(masterFilePath));
            string tempUnifiedMasterFile = Path.Combine(tempWorkFolder, Path.GetFileName(unifiedMasterFilePath));
            File.Copy(masterFilePath, tempMasterFile, true);
            File.Copy(unifiedMasterFilePath, tempUnifiedMasterFile, true);

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
            Workbook unifiedMasterFile = null;
            Worksheet worksheet = null;
            Worksheet unifiedWorksheet = null;
            Range usedRange = null;
            Range unifiedUsedRange = null;

            try
            {
                excelApp = new Application();
                masterFile = excelApp.Workbooks.Open(tempMasterFile);
                unifiedMasterFile = excelApp.Workbooks.Open(tempUnifiedMasterFile);
                worksheet = (Worksheet?)masterFile.Worksheets[1];
                unifiedWorksheet = (Worksheet?)unifiedMasterFile.Worksheets[1];
                usedRange = worksheet.UsedRange;
                unifiedUsedRange = unifiedWorksheet.UsedRange;

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
                            await ProcessBatchAsync(batch, excelApp, maxRows, tempWorkFolder, tempOutputDir, unifiedWorksheet);
                            batch.Clear();
                            GC.Collect();
                        }
                    }

                    // Process any remaining rows
                    if (batch.Count > 0)
                    {
                        await ProcessBatchAsync(batch, excelApp, maxRows, tempWorkFolder, tempOutputDir, unifiedWorksheet);
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

        private async Task ProcessBatchAsync(List<Dictionary<string, string>> batch, Application excelApp, int numberRowsToGenerate, string tempWorkFolder, string tempOutputDir, Worksheet unifiedMasterDataSheet = null)
        {
            foreach (var row in batch)
            {
                await ProcessRow(row, excelApp, numberRowsToGenerate, tempWorkFolder, tempOutputDir, unifiedMasterDataSheet);
            }
        }

        private async Task ProcessRow(Dictionary<string, string> row, Application excelApp, int numberRowsToGenerate, string tempWorkFolder, string tempOutputDir, Worksheet unifiedWorksheet = null)
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
            Worksheet mainSheet = null;
            Worksheet staffListSheet = null;

            var safeSubjectCode = string.Concat(subjectCode.Split(Path.GetInvalidFileNameChars())).ToLowerInvariant();
            var safeStudyPeriod = string.Concat(studyPeriod.Split(Path.GetInvalidFileNameChars())).ToLowerInvariant();
            var identifier = $"{safeSubjectCode}_{safeStudyPeriod}";
            var fileName = $"{identifier}.xlsx";

            Range staffListRange = null;
            List<int> foundRows = new List<int>();
            if (unifiedWorksheet != null)
            {
                foundRows = FindMatchingRowsUnifiedMaster(unifiedWorksheet, subjectCode, studyPeriod);
            }

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

                if (foundRows.Count > 0)
                {
                    mainSheet = (Worksheet)workbook.Worksheets[1];
                    mainSheet.Unprotect(ConfigManager.Template_File_Password);
                    staffListSheet = (Worksheet)workbook.Worksheets[ConfigManager.StaffList_Sheet_Name];
                    var usedUnifiedMasterRange = unifiedWorksheet.UsedRange;

                    if (staffListSheet != null)
                    {
                        staffListRange = staffListSheet.Columns[5];
                        staffListSheet.Unprotect(ConfigManager.Template_File_Password);
                    }

                    FindStaffRanges(mainSheet, out int startRow, out int endRow, out int otherStaffStartRow, out int otherStaffEndRow);

                    //copy the filter range into dataSheet starting from B24

                    //row 1 is header, so start from row 2
                    for (int r = 0; r < foundRows.Count; r++)
                    {
                        var v = GetCellValueAsString(usedUnifiedMasterRange.Cells[foundRows[r], 1]);
                        //Staff name in column 6
                        string staffName = usedUnifiedMasterRange.Cells[foundRows[r], 6]?.Value2?.ToString();
                        if (string.IsNullOrWhiteSpace(staffName))
                        {
                            continue; // Skip rows where column F is empty
                        }
                        else
                        {
                            if (string.Equals(staffName.Trim(), ConfigManager.Non_Safes_UoM_Staff, StringComparison.OrdinalIgnoreCase))
                            {
                                for (int c = 3; c <= 12; c++)
                                {
                                    mainSheet.Cells[otherStaffStartRow, c] = usedUnifiedMasterRange.Cells[foundRows[r], c + 4]?.Value2;
                                }
                            }
                            else if (string.Equals(staffName.Trim(), ConfigManager.Casual_Lecturers, StringComparison.OrdinalIgnoreCase))
                            {
                                for (int c = 3; c <= 12; c++)
                                {
                                    mainSheet.Cells[otherStaffStartRow + 1, c] = usedUnifiedMasterRange.Cells[foundRows[r], c + 4]?.Value2;
                                }
                            }
                            else if (string.Equals(staffName.Trim(), ConfigManager.Casual_Tutors, StringComparison.OrdinalIgnoreCase))
                            {
                                for (int c = 3; c <= 12; c++)
                                {
                                    mainSheet.Cells[otherStaffStartRow + 2, c] = usedUnifiedMasterRange.Cells[foundRows[r], c + 4]?.Value2;
                                }
                            }
                            else
                            {
                                //Find staffname in staff list cloumn 5, from row 4
                                //if not exist then insert the staff name into staff list 
                                if (staffListRange != null)
                                {
                                    Range found = staffListRange.Find(staffName, LookIn: XlFindLookIn.xlValues, LookAt: XlLookAt.xlWhole);
                                    if (found == null)
                                    {
                                        int newRow = FindFirstEmptyRowInColumn(staffListSheet, 5, 4);
                                        staffListSheet.Cells[newRow, 5] = staffName;


                                        // Create the data validation formula
                                        for (int i = startRow; i <= endRow; i++)
                                        {
                                            Range cell = mainSheet.Cells[i, 2] as Range;
                                            if (cell != null)
                                            {
                                                try
                                                {
                                                    string formula = $"='{ConfigManager.StaffList_Sheet_Name}'!$E$4:$E${newRow + 3}";
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
                                                catch (Exception ex)
                                                {
                                                    Debug.WriteLine($"Failed to set data validation for cell {cell.Address}: {ex.Message}");
                                                }
                                            }
                                        }
                                    }
                                }

                                mainSheet.Cells[startRow, 2] = staffName;
                                for (int c = 3; c <= 12; c++)
                                {
                                    try
                                    {
                                        string val = GetCellValueAsString(usedUnifiedMasterRange.Cells[foundRows[r], c + 4]);
                                        mainSheet.Cells[startRow, c] = val;
                                    }
                                    catch (Exception ex)
                                    {
                                        Debug.WriteLine($"Failed to write value for staff {staffName} at main sheet row {startRow}, column {c}: {ex.Message}");
                                    }
                                }
                                startRow++;
                            }
                        }

                    }

                }
                //Remove filter
                unifiedWorksheet.AutoFilterMode = false;


                //Protect the workbook again
                mainSheet?.Protect(ConfigManager.Template_File_Password);
                staffListSheet?.Protect(ConfigManager.Template_File_Password);
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
                if (mainSheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(mainSheet);
                if (staffListSheet != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(staffListSheet);
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
                    "Practical Initial", "Practical Repeat", "FieldTrip/Excursion Lead", "FieldTrip/Excursion Assisting", "Marking"
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

                            FindStaffRanges(srcWs, out startRow, out endRow, out otherStaffStartRow, out otherStaffEndRow);

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
            else if (btnViewOutput.Text.Contains("Staff Summary"))
            {
                folderPath = ConfigManager.Staff_Summary_Output_Location;
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

        private void btnViewUnifiedMaster_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe", Path.GetDirectoryName(txtUnifiedMasterFile.Text));
        }

        // Add this helper inside the Main class (near other private helpers)
        private void FindStaffRanges(Worksheet srcWs, out int startRow, out int endRow, out int otherStaffStartRow, out int otherStaffEndRow)
        {
            startRow = 0;
            endRow = 0;
            otherStaffStartRow = 0;
            otherStaffEndRow = 0;

            try
            {
                int totalRows = srcWs.UsedRange.Rows.Count;

                // Find SafesStaff start and end
                for (int row = 1; row <= totalRows; row++)
                {
                    var cellValue = (srcWs.Cells[row, 2] as Range)?.Value2?.ToString();
                    if (string.IsNullOrWhiteSpace(cellValue)) continue;

                    var txt = cellValue.Trim();
                    if (string.Equals(txt, ConfigManager.SafesStaff_Label.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        startRow = row + 2; // Start row is the row after the label
                    }
                    else if (txt.IndexOf(ConfigManager.TotalHrs_Label, StringComparison.OrdinalIgnoreCase) >= 0
                             && startRow > 0 && row > startRow)
                    {
                        endRow = row - 1; // End row is the row before this label
                        break;
                    }
                }

                // Find OtherStaff start and end (search after endRow if available)
                int searchStart = endRow > 0 ? endRow + 1 : 1;
                for (int row = searchStart; row <= totalRows; row++)
                {
                    var cellValue = (srcWs.Cells[row, 2] as Range)?.Value2?.ToString();
                    if (string.IsNullOrWhiteSpace(cellValue)) continue;

                    var txt = cellValue.Trim();
                    if (string.Equals(txt, ConfigManager.OtherStaff_Label.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        otherStaffStartRow = row + 2; // Start row is the row after the label
                    }
                    else if (txt.IndexOf(ConfigManager.TotalHrs_Label, StringComparison.OrdinalIgnoreCase) >= 0
                             && otherStaffStartRow > 0
                             && otherStaffStartRow > startRow
                             && row > otherStaffStartRow)
                    {
                        otherStaffEndRow = row - 1; // End row is the row before this label
                        break;
                    }
                }
            }
            catch (COMException comEx)
            {
                Debug.WriteLine($"FindStaffRanges COM error: {comEx.Message}");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"FindStaffRanges error: {ex.Message}");
            }
        }

        private int FindFirstEmptyRowInColumn(Worksheet ws, int columnIndex, int startRow)
        {
            if (ws == null) throw new ArgumentNullException(nameof(ws));
            Range usedRange = null;
            try
            {
                usedRange = ws.UsedRange;
                int usedFirstRow = usedRange.Row;
                int usedLastRow = usedFirstRow + usedRange.Rows.Count - 1;
                int scanStart = Math.Max(startRow, usedFirstRow);

                // Scan within the used area for the first empty cell in the column
                for (int r = scanStart; r <= usedLastRow; r++)
                {
                    Range cell = null;
                    try
                    {
                        cell = ws.Cells[r, columnIndex] as Range;
                        var value = cell?.Value2;
                        if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
                        {
                            return r; // first empty row found inside used area
                        }
                    }
                    finally
                    {
                        if (cell != null) Marshal.ReleaseComObject(cell);
                    }
                }

                // No empty row inside used area — return next row after used area
                return usedLastRow + 1;
            }
            finally
            {
                if (usedRange != null) Marshal.ReleaseComObject(usedRange);
            }
        }

        private static string GetCellValueAsString(Range? cell)
        {
            if (cell == null) return string.Empty;

            try
            {
                var val = cell.Value2;
                if (val == null) return string.Empty;

                // Strings
                if (val is string s) return s.Trim();

                // Booleans
                if (val is bool b) return b ? "TRUE" : "FALSE";

                // Numbers (including Excel dates as OLE Automation dates)
                if (val is double d)
                {
                    // Try detect if the cell is formatted as a date
                    try
                    {
                        var nf = (cell.NumberFormat as string) ?? string.Empty;
                        var nfLower = nf.ToLowerInvariant();
                        // crude check for date number formats (adjust if needed)
                        if (nfLower.Contains("d") || nfLower.Contains("y") || nfLower.Contains("/") || nfLower.Contains("-"))
                        {
                            // treat as OLE date
                            try
                            {
                                var dt = DateTime.FromOADate(d);
                                return dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture).TrimEnd(' ', '0', ':');
                            }
                            catch
                            {
                                // fallback to numeric formatting if conversion fails
                            }
                        }
                    }
                    catch
                    {
                        // ignore NumberFormat read errors and fall back to numeric formatting
                    }

                    // numeric value
                    return d.ToString(CultureInfo.InvariantCulture);
                }

                // Other types fallback
                return val.ToString() ?? string.Empty;
            }
            finally
            {
                // do not release `cell` here - caller should release COM objects when appropriate
            }
        }

        // Add inside the Main class near other private helpers
        public static List<int> FindMatchingRowsUnifiedMaster(
            Worksheet worksheet,
            string subjectCode,
            string studyPeriod)
        {
            Range used = worksheet.UsedRange;

            int rowCount = used.Rows.Count;

            List<int> matchedRows = new List<int>();

            string subjectTrim = subjectCode?.Trim();
            string periodTrim = studyPeriod?.Trim();

            for (int r = 2; r <= rowCount; r++) // assume header row = 1
            {
                var subject = (used.Cells[r, 1] as Range)?.Value2?.ToString()?.Trim();
                var period = (used.Cells[r, 3] as Range)?.Value2?.ToString()?.Trim();

                if (subject == subjectTrim &&
                    period == periodTrim)
                {
                    matchedRows.Add(r);
                }
            }

            return matchedRows;
        }

        private async void btnStaffSummaryGenerate_Click(object sender, EventArgs e)
        {
            lstReport.Items.Clear();
            Cursor = Cursors.WaitCursor;
            progressBar.MarqueeAnimationSpeed = 30;

            Invoke(new System.Action(() =>
            {
                lblActionDisplay.Visible = true;
                lblActionDisplay.Text = "Create Staff Summary files...";
                progressBar.Visible = true;
                progressBar.Style = ProgressBarStyle.Marquee;
                lstReport.Visible = true;
            }));

            //check if the unified data exists
            if (!File.Exists(txtMasterDashboard.Text))
            {
                MessageBox.Show("Master Dashboard file does not exist.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Cursor = Cursors.Default;
                progressBar.Style = ProgressBarStyle.Blocks;
                progressBar.Visible = false;
                return;
            }

            var templatePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Data", "staff-summary.xlsx");
            if (!File.Exists(templatePath))
            {
                MessageBox.Show($"Template not found: {templatePath}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Cursor = Cursors.Default;
                progressBar.Style = ProgressBarStyle.Blocks;
                progressBar.Visible = false;
                return;
            }

            //create temp folder for staff summary generation
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            string tempStaffSummaryFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Working", $"StaffSummary_TempWork_{timestamp}");
            //copy the unified data file to the temp folder
            Directory.CreateDirectory(tempStaffSummaryFolder);
            
            string masterDashboardFile = Path.Combine(tempStaffSummaryFolder, Path.GetFileName(txtMasterDashboard.Text));
            File.Copy(txtMasterDashboard.Text, masterDashboardFile, true);

            //create a folder to store output files 
            string outputFolder = Path.Combine(tempStaffSummaryFolder, "Output");
            Directory.CreateDirectory(outputFolder);

            // Kill all running Excel processes before starting
            foreach (var process in System.Diagnostics.Process.GetProcessesByName("EXCEL"))
            {
                try { process.Kill(); }
                catch { /* ignore if cannot kill */ }
            }

            Application excelApp = null;
            //Workbook unifiedWb = null;
            Workbook masterDashBoardWb = null;
            Worksheet ws = null;
            Worksheet mDbWs = null;
            Range used = null;

            try
            {
                excelApp = new Application();
                bool unprotected;
                //unifiedWb = OpenWorkbookSilent(excelApp, destUnifiedDataFile, ConfigManager.Template_File_Password, ConfigManager.Template_File_Password);//OpenWorkbookAndEnsureUnprotected(excelApp, destUnifiedDataFile, ConfigManager.Template_File_Password, out unprotected); //excelApp.Workbooks.Open(destUnifiedDataFile);
                masterDashBoardWb = OpenWorkbookSilent(excelApp, masterDashboardFile, ConfigManager.Template_File_Password, ConfigManager.Template_File_Password);//OpenWorkbookAndEnsureUnprotected(excelApp, masterDashboardFile, ConfigManager.Template_File_Password, out unprotected);//excelApp.Workbooks.Open(masterDashboardFile);

                ws = masterDashBoardWb.Worksheets["Unify data"] as Worksheet; //unifiedWb.Worksheets[1] as Worksheet;
                mDbWs = masterDashBoardWb.Worksheets[1] as Worksheet;

                if (ws == null)
                {
                    MessageBox.Show("No 'Unify Data' sheet in Master Dashboard file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Cursor = Cursors.Default;
                    progressBar.Style = ProgressBarStyle.Blocks;
                    progressBar.Visible = false;
                    return;
                }

                //Create all files for each staff in this mDbWs first
                try
                {
                    Range staffRange = mDbWs.UsedRange;
                    await Task.Run(async () =>
                    {

                        for (int r = 2; r <= staffRange.Rows.Count; r++)
                        {

                            Workbook staffWb = null;
                            Worksheet staffWs = null;
                            Worksheet staffDataWs = null;


                            string staffName = GetCellValueAsString(staffRange.Cells[r, 1]); // column A
                            if (string.IsNullOrWhiteSpace(staffName))
                            {
                                //continue; // skip empty staff
                                //empty row - should be the end of data - break the loop
                                break;
                            }
                            try
                            {


                                staffName = staffName.Trim();
                                var safeName = SanitizeFileName(staffName);

                                Invoke(new System.Action(() =>
                                {
                                    lstReport.Items.Add($"Created: {safeName}.xlsx");
                                    lstReport.SelectedIndex = lstReport.Items.Count - 1;
                                }));

                                string outFile = Path.Combine(outputFolder, $"{safeName}.xlsx");
                                // if file exists already append index
                                int idx = 1;
                                while (File.Exists(outFile))
                                {
                                    outFile = Path.Combine(outputFolder, $"{safeName}_{idx}.xlsx");
                                    idx++;
                                }

                                File.Copy(templatePath, outFile, overwrite: true);

                                staffWb = excelApp.Workbooks.Open(outFile);
                                staffWs = staffWb.Worksheets[1] as Worksheet;
                                staffDataWs = staffWb.Worksheets["Data"] as Worksheet;

                                staffWs.Cells[2, 3] = staffName; // Put staff name into cell C2 (column 3, row 2)
                                                                 //clear the existing data in outDataWs before writing new data - row 2 onwards
                                Range clearRange = staffDataWs.Range["A2:Y2"]; // Adjust the range as needed
                                clearRange.ClearContents();

                                // Copy the row to staffDataWs starting at row 2
                                for (int c = 1; c <= staffRange.Columns.Count; c++)
                                {
                                    var val = GetCellValueAsString(staffRange.Cells[r, c]);
                                    staffDataWs.Cells[2, c] = val;
                                }

                                staffDataWs.Visible = XlSheetVisibility.xlSheetVeryHidden;

                                staffWb.Save();
                                staffWb.Close();

                            }
                            catch { }
                            finally
                            {
                                if (staffWs != null) Marshal.ReleaseComObject(staffWs);
                                if (staffDataWs != null) Marshal.ReleaseComObject(staffDataWs);
                                if (staffWb != null)
                                {
                                    try { Marshal.ReleaseComObject(staffWb); }
                                    catch { }
                                }
                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                            }
                        }
                    });
                }
                catch
                { }


                lstReport.Items.Clear();
                Invoke(new System.Action(() =>
                {
                    lblActionDisplay.Visible = true;
                    lblActionDisplay.Text = "Reading Unify Data and update Staff Summary file...";
                    progressBar.Visible = true;
                    progressBar.Style = ProgressBarStyle.Marquee;
                    lstReport.Visible = true;
                }));
               
                //Build template Data - End

                //Start building subject rows
                used = ws.UsedRange;

                int totalRows = used.Rows.Count;
                int totalCols = used.Columns.Count;

                if (totalRows < 2)
                {
                    MessageBox.Show("Unified file does not contain data rows.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Read header row (row 1)
                var headers = new List<string>(totalCols);
                for (int c = 1; c <= totalCols; c++)
                {
                    headers.Add(GetCellValueAsString(used.Cells[1, c]));
                }

                // Group rows by staff name (column F = 6)
                var staffGroups = new Dictionary<string, List<int>>(StringComparer.OrdinalIgnoreCase);
                for (int r = 2; r <= totalRows; r++)
                {
                    string staffName = GetCellValueAsString(used.Cells[r, 8]); // column H
                    if (string.IsNullOrWhiteSpace(staffName))
                    {
                        //continue; // skip empty staff
                        //empty row - should be the end of data - break the loop
                        break;
                    }

                    staffName = staffName.Trim();
                    if (!staffGroups.TryGetValue(staffName, out var list))
                    {
                        list = new List<int>();
                        staffGroups[staffName] = list;
                    }
                    list.Add(r);
                }

                int processedStaff = 0;
                int totalStaff = staffGroups.Count;

                // Mapping: unified columns -> we will populate template columns starting at B9:
                // B = Combine (Subject Code (unified col 1) and Study Period (unified col 3) combined)
                // C = Coordinator (unified col 10)
                // D = Lecture Initial (unified col 11)
                // E = Lecture Repeat  (unified col 12)
                // F = Tute/WS Initial (unified col 13)
                // G = Tute/WS Repeat  (unified col 14)
                // H = Practical Initial (unified col 15)
                // I = Practical Repeat  (unified col 16)
                // J = FieldTrip Lead     (unified col 17)
                // K = FieldTrip Assisting(unified col 18)
                // L = Marking            (unified col 19)
                int[] unifiedColsForTemplate = new int[] { 10, 11, 12, 13, 14, 15, 16, 17, 18, 19 };

                await Task.Run(async () =>
                {
                    // Create one file per staff
                    foreach (var kvp in staffGroups)
                    {
                        string staffName = kvp.Key;

                        var rows = kvp.Value;

                        // sanitize file name - need to lowcase and replace space or special characters to _
                        var safeName = SanitizeFileName(staffName);

                        string outFile = Path.Combine(outputFolder, $"{safeName}.xlsx");

                        Workbook outWb = null;
                        Worksheet outWs = null;
                        Worksheet outDataWs = null;

                        try
                        {
                            bool createFile = false;
                            // copy the template for this staff (preserves formatting, hidden sheets, etc.)
                            if (!File.Exists(outFile))
                            {
                                createFile = true;
                                File.Copy(templatePath, outFile, overwrite: true);
                            }
                            outWb = excelApp.Workbooks.Open(outFile);
                            outWs = outWb.Worksheets[1] as Worksheet;
                            outDataWs = outWb.Worksheets["Data"] as Worksheet;


                            //we need to clear the existing data in the template before writing new data
                            Range dataRange = outWs.Range["B9:L1000"]; // Adjust the range as needed
                            dataRange.ClearContents();

                            // Start writing subject rows at row 9 (B9 is subject code)
                            int writeRow = 9;
                            foreach (var srcRow in rows)
                            {
                                // With this combined Subject Code + Study Period value:
                                var subjectCode = GetCellValueAsString(used.Cells[srcRow, 1]);
                                var studyPeriod = GetCellValueAsString(used.Cells[srcRow, 3]);
                                outWs.Cells[writeRow, 2] = string.IsNullOrWhiteSpace(studyPeriod)
                                    ? subjectCode.Trim()
                                    : $"{subjectCode.Trim()} - {studyPeriod.Trim()}";

                                // Fill coordinator / lecture / tute / practical / fieldtrip / marking starting at column C (col 3)
                                for (int i = 0; i < unifiedColsForTemplate.Length; i++)
                                {
                                    int unifiedCol = unifiedColsForTemplate[i];
                                    int targetCol = 3 + i; // 3 => C, 4 => D, ...
                                    string val = GetCellValueAsString(used.Cells[srcRow, unifiedCol]);
                                    outWs.Cells[writeRow, targetCol] = val;
                                }

                                writeRow++;
                            }

                            // Optional: Autofit only the populated columns to preserve template formatting elsewhere
                            //Range dataRange = outWs.Range[outWs.Cells[9, 2], outWs.Cells[Math.Max(9, writeRow - 1), 12]]; // B..L
                            //dataRange.Columns.AutoFit();
                            //Marshal.ReleaseComObject(dataRange);

                            // Save and close
                            outWb.Save();
                            outWb.Close(false);

                            processedStaff++;
                            string displayText = $"[{processedStaff}/{totalStaff}] Update: {Path.GetFileName(outFile)} ({rows.Count} rows)";
                            if (createFile)
                            {
                                displayText = $"[{processedStaff}/{totalStaff}] Create: {Path.GetFileName(outFile)} ({rows.Count} rows)";
                            }
                            Invoke(new System.Action(() =>
                            {
                                lstReport.Items.Add(displayText);
                                lblReport.Text = $"Update {processedStaff} of {totalStaff} staff summary files...";
                                lstReport.SelectedIndex = lstReport.Items.Count - 1;
                            }));

                        }
                        catch (Exception ex)
                        {
                            Invoke(new System.Action(() =>
                            {
                                lstReport.Items.Add($"Error creating file for '{staffName}': {ex.Message}");
                            }));
                        }
                        finally
                        {
                            if (outWs != null) Marshal.ReleaseComObject(outWs);
                            if (outWb != null)
                            {
                                try { Marshal.ReleaseComObject(outWb); }
                                catch { }
                            }
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                    }
                }); // End of Task.Run

                //copy all files from output folder to Staff Summary Output Location overwrite existing files
                string sharepointStaffSummaryFolder = ConfigManager.Staff_Summary_Output_Location;

                if (!Directory.Exists(sharepointStaffSummaryFolder))
                {
                    Directory.CreateDirectory(sharepointStaffSummaryFolder);
                }
                else
                {
                    //Copy overwrite - copy all files from output folder to Staff Summary Output Location overwrite existing files
                    foreach (var file in Directory.GetFiles(outputFolder, "*.xlsx", SearchOption.TopDirectoryOnly))
                    {
                        var destFile = Path.Combine(sharepointStaffSummaryFolder, Path.GetFileName(file));
                        File.Copy(file, destFile, true); // true = overwrite existing files
                    }
                }
                // Final UI update
                Invoke(new System.Action(() =>
                {
                    lblActionDisplay.Text = "Staff summary generation completed.";
                    btnViewOutput.Visible = true;
                    btnViewOutput.Text = "View Staff Summary Folder";
                }));
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating staff summaries: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (used != null) Marshal.ReleaseComObject(used);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (mDbWs != null) Marshal.ReleaseComObject(mDbWs);
                
                if (masterDashBoardWb != null)
                {
                    masterDashBoardWb.Close(false);
                    Marshal.ReleaseComObject(masterDashBoardWb);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();

                Invoke(new System.Action(() =>
                {
                    Cursor = Cursors.Default;
                    progressBar.Style = ProgressBarStyle.Blocks;
                    progressBar.Visible = false;
                }));
            }
        }

        private static string SanitizeFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "unknownstaff";

            // lowercase and trim
            var s = name.ToLowerInvariant().Trim();

            // Replace any invalid filesystem characters with underscore
            foreach (var c in Path.GetInvalidFileNameChars())
            {
                s = s.Replace(c, '_');
            }

            // Replace any remaining characters that are not a-z, 0-9, dot or dash with underscore
            s = System.Text.RegularExpressions.Regex.Replace(s, @"[^a-z0-9\.\-]+", "_");

            // Collapse multiple underscores into a single underscore
            s = System.Text.RegularExpressions.Regex.Replace(s, "_{2,}", "_");

            // Trim leading/trailing underscores, dots or dashes
            s = s.Trim('_', '.', '-');

            return string.IsNullOrEmpty(s) ? "unknownstaff" : s;
        }

        private void btnViewMasterDashboardData_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe", Path.GetDirectoryName(txtMasterDashboard.Text));
        }

        // Add inside the `Main` class near other private helpers
        private Workbook OpenWorkbookAndEnsureUnprotected(Application excelApp, string filePath, string password, out bool anyUnprotected)
        {
            if (excelApp == null) throw new ArgumentNullException(nameof(excelApp));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));

            anyUnprotected = false;
            Workbook wb = null;

            // Open workbook (read/write)
            wb = excelApp.Workbooks.Open(filePath);

            try
            {
                // Attempt to unprotect workbook-level protection (structure/windows).
                // Unprotect will succeed silently if there is no protection, or throw if password is incorrect.
                try
                {
                    wb.Unprotect(password);
                    // If no exception, treat as unprotected or already unprotected
                    anyUnprotected = true;
                }
                catch (COMException)
                {
                    // If wrong password or other COM error, we still continue to try unprotecting sheets.
                }

                // Iterate worksheets and unprotect any protected sheets
                foreach (Worksheet ws in wb.Worksheets)
                {
                    try
                    {
                        bool isProtected = false;
                        try
                        {
                            // Check common protection flags
                            isProtected = (ws.ProtectContents || ws.ProtectDrawingObjects || ws.ProtectScenarios);
                        }
                        catch
                        {
                            // If any COM error reading flags, assume possibly protected and try to unprotect
                            isProtected = true;
                        }

                        if (isProtected)
                        {
                            try
                            {
                                ws.Unprotect(password);
                                anyUnprotected = true;
                            }
                            catch (COMException)
                            {
                                // ignore - wrong password or cannot unprotect; continue
                            }
                        }
                    }
                    finally
                    {
                        // release worksheet COM object
                        if (ws != null) Marshal.ReleaseComObject(ws);
                    }
                }

                return wb;
            }
            catch
            {
                // if we fail here, ensure the workbook is closed and released
                if (wb != null)
                {
                    try { wb.Close(false); } catch { }
                    Marshal.ReleaseComObject(wb);
                }
                throw;
            }
        }

        private int FindFirstRowByStaffName(Worksheet ws, string staffName, int staffColumn = 1, int startRow = 2)
        {
            if (ws == null) throw new ArgumentNullException(nameof(ws));
            if (string.IsNullOrWhiteSpace(staffName)) return -1;

            // Simple scan: start at `startRow`, stop when cell is empty or when a match is found.
            // from row 2 and until found or empty then exit".
            const int ExcelMaxRows = 1_048_576;
            var target = staffName.Trim();

            for (int r = startRow; r <= ExcelMaxRows; r++)
            {
                Range cell = null;
                try
                {
                    cell = ws.Cells[r, staffColumn] as Range;
                    var cellValue = GetCellValueAsString(cell);

                    // stop when we hit the first empty cell in the staff column
                    if (string.IsNullOrWhiteSpace(cellValue))
                    {
                        break;
                    }

                    if (string.Equals(cellValue.Trim(), target, StringComparison.OrdinalIgnoreCase))
                    {
                        return r;
                    }
                }
                finally
                {
                    if (cell != null) Marshal.ReleaseComObject(cell);
                }
            }

            return -1;
        }

        // Add inside the `Main` class near other private helpers

        /// <summary>
        /// Open a workbook without user prompts by supplying open/write passwords and disabling alerts.
        /// Attempts to unprotect workbook and sheets using the supplied password (if any).
        /// </summary>
        private Workbook OpenWorkbookSilent(Application excelApp, string filePath, string openPassword = null, string writePassword = null)
        {
            if (excelApp == null) throw new ArgumentNullException(nameof(excelApp));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentNullException(nameof(filePath));

            // Suppress prompts
            var prevDisplayAlerts = excelApp.DisplayAlerts;
            try
            {
                excelApp.DisplayAlerts = false;
                excelApp.AskToUpdateLinks = false;

                object oFilename = filePath;
                object oUpdateLinks = 0;
                object oReadOnly = false;
                object oFormat = Type.Missing;
                object oPassword = string.IsNullOrEmpty(openPassword) ? Type.Missing : (object)openPassword;
                object oWriteResPassword = string.IsNullOrEmpty(writePassword) ? Type.Missing : (object)writePassword;
                object oIgnoreReadOnlyRecommended = true;
                object oOrigin = Type.Missing;
                object oDelimiter = Type.Missing;
                object oEditable = true;
                object oNotify = false;
                object oConverter = Type.Missing;
                object oAddToMru = false;
                object oLocal = true;
                object oCorruptLoad = Type.Missing;

                Workbook wb = excelApp.Workbooks.Open(
                    Filename: (string)oFilename,
                    UpdateLinks: oUpdateLinks,
                    ReadOnly: oReadOnly,
                    Format: oFormat,
                    Password: oPassword,
                    WriteResPassword: oWriteResPassword,
                    IgnoreReadOnlyRecommended: oIgnoreReadOnlyRecommended,
                    Origin: oOrigin,
                    Delimiter: oDelimiter,
                    Editable: oEditable,
                    Notify: oNotify,
                    Converter: oConverter,
                    AddToMru: oAddToMru,
                    Local: oLocal,
                    CorruptLoad: oCorruptLoad
                );

                // Try unprotecting workbook structure and sheets silently (ignore failures)
                try { wb.Unprotect(openPassword ?? writePassword ?? string.Empty); } catch { /* ignore */ }

                foreach (Worksheet ws in wb.Worksheets)
                {
                    try { ws.Unprotect(openPassword ?? writePassword ?? string.Empty); }
                    catch { /* ignore */ }
                    finally { if (ws != null) Marshal.ReleaseComObject(ws); }
                }

                return wb;
            }
            finally
            {
                // Restore DisplayAlerts in the caller when appropriate; we restore here to be safe.
                excelApp.DisplayAlerts = prevDisplayAlerts;
            }
        }

    }
}
