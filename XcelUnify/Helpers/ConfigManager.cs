using System.Diagnostics;
using System.Text.Json;


namespace XcelUnify.Helpers
{
    public static class ConfigManager
    {
        public static string Master_File { get; private set; } = string.Empty;
        public static string Template_File { get; private set; } = string.Empty;
        public static string Output_Location { get; private set; } = string.Empty;
        public static string Unify_Folder { get; private set; } = string.Empty;
        public static string Done_Folder_Format { get; private set; } = "yyyyMMddHHmm_Done";
        public static string Report_File_Format { get; private set; } = "yyyyMMddHHmm_UnifyRpt.xlsx";
        public static int Workload_Main_Sheet { get; private set; } = 2;
        public static string Coursework_Text { get; private set; } = "N";
        public static int Max_Rows { get; private set; } = 0;
        public static int Batch_Size { get; private set; } = 10;
        public static int SafesStaff_Start_Row { get; private set; } = 22;
        public static int SafesStaff_End_Row { get; private set; } = 30;

        public static int OtherStaff_Start_Row { get; private set; } = 22;
        public static int OtherStaff_End_Row { get; private set; } = 30;


        public static void Init()
        {
            try
            {
                using var stream = File.OpenRead(Path.Combine(AppContext.BaseDirectory, "appsettings.json"));
                using var reader = new StreamReader(stream);
                var configText = reader.ReadToEnd();

                var config = JsonSerializer.Deserialize<Dictionary<string, string>>(configText);

                if (config != null)
                {
                    config.TryGetValue("MasterDataFilePath", out var masterFile);
                    config.TryGetValue("TemplateFilePath", out var templateFile);
                    config.TryGetValue("Output", out var output);
                    config.TryGetValue("Unify", out var unifyFolder);
                    config.TryGetValue("DoneFolderFormat", out var doneFolderFormat);
                    config.TryGetValue("ReportFileFormat", out var reportFileFormat);
                    config.TryGetValue("WorkloadMainSheet", out var workloadMainSheet);
                    config.TryGetValue("CourseworkText", out var courseworkText);
                    config.TryGetValue("MaxRowsToGenerate", out var maxRows);
                    config.TryGetValue("BatchSizeToGenerate", out var batchSize);
                    config.TryGetValue("SafesStaffStartRow", out var safesStaffStartRow);
                    config.TryGetValue("SafesStaffEndRow", out var safesStaffEndRow);
                    config.TryGetValue("OtherStaffStartRow", out var otherStaffStartRow);
                    config.TryGetValue("OtherStaffEndRow", out var otherStaffEndRow);

                    Master_File = !string.IsNullOrEmpty(masterFile)
                        ? Path.Combine(AppContext.BaseDirectory, masterFile)
                        : string.Empty;

                    Template_File = !string.IsNullOrEmpty(templateFile)
                        ? Path.Combine(AppContext.BaseDirectory, templateFile)
                        : string.Empty;

                    Output_Location = !string.IsNullOrEmpty(output)
                        ? Path.Combine(AppContext.BaseDirectory, output)
                        : string.Empty;

                    Unify_Folder = !string.IsNullOrEmpty(unifyFolder)
                        ? Path.Combine(AppContext.BaseDirectory, unifyFolder)
                        : string.Empty;

                    Done_Folder_Format = !string.IsNullOrEmpty(doneFolderFormat)
                        ? doneFolderFormat
                        : "yyyyMMddHHmm_Done";

                    Report_File_Format = !string.IsNullOrEmpty(reportFileFormat)
                        ? reportFileFormat
                        : "yyyyMMddHHmm_UnifyRpt.xlsx";

                    Workload_Main_Sheet = !string.IsNullOrEmpty(workloadMainSheet) ?
                        System.Convert.ToInt32(workloadMainSheet) : 2;

                    Coursework_Text = !string.IsNullOrEmpty(courseworkText) ? courseworkText : string.Empty;

                    Max_Rows = !string.IsNullOrEmpty(maxRows) ?
                        System.Convert.ToInt32(maxRows) : 0;
                    Batch_Size = !string.IsNullOrEmpty(batchSize) ?
                        System.Convert.ToInt32(batchSize) : 10;

                    SafesStaff_Start_Row = !string.IsNullOrEmpty(safesStaffStartRow) ?
                        System.Convert.ToInt32(safesStaffStartRow) : 22;

                    SafesStaff_End_Row = !string.IsNullOrEmpty(safesStaffEndRow) ?
                        System.Convert.ToInt32(safesStaffEndRow) : 30;

                    OtherStaff_Start_Row = !string.IsNullOrEmpty(otherStaffStartRow) ?
                        System.Convert.ToInt32(otherStaffStartRow) : 35;

                    OtherStaff_End_Row = !string.IsNullOrEmpty(otherStaffEndRow) ?
                        System.Convert.ToInt32(otherStaffEndRow) : 37;
                }
                else
                {
                    Master_File = string.Empty;
                    Template_File = string.Empty;
                    Output_Location = string.Empty;
                    Unify_Folder = string.Empty;
                    Done_Folder_Format = "yyyyMMddHHmm_Done";
                    Report_File_Format = "yyyyMMddHHmm_UnifyRpt.xlsx";
                    Coursework_Text = string.Empty;
                    SafesStaff_Start_Row = 22;
                    SafesStaff_End_Row = 30;
                    OtherStaff_Start_Row = 35;
                    OtherStaff_End_Row = 37;
                }
            }
            catch
            {
                Master_File = string.Empty;
                Template_File = string.Empty;
                Output_Location = string.Empty;
                Unify_Folder = string.Empty;
                Done_Folder_Format = "yyyyMMddHHmm_Done";
                Report_File_Format = "yyyyMMddHHmm_UnifyRpt.xlsx";
                Coursework_Text = string.Empty;
                SafesStaff_Start_Row = 22;
                SafesStaff_End_Row = 30;
                OtherStaff_Start_Row = 35;
                OtherStaff_End_Row = 37;
            }
        }

    }
}
