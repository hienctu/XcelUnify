using Microsoft.Office.Interop.Excel;
using XcelUnify.Helpers;
using Range = Microsoft.Office.Interop.Excel.Range;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace XcelUnify
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private async void btnGenerate_Click(object sender, EventArgs e)
        {
            int maxRows = 5;
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

                // Data rows - skip header row
                const int batchSize = 2;
                var batch = new List<Dictionary<string, string>>(batchSize);
                int processed = 0;

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

                // Copy template to writable location
                File.Copy(ConfigManager.Template_File, targetPath, overwrite: true);

                // Open the copied template file
                workbook = excelApp.Workbooks.Open(targetPath);

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
    }
}
