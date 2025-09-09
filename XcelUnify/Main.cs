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
    }
}
