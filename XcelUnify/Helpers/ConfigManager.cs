using System.Diagnostics;
using System.Text.Json;


namespace XcelUnify.Helpers
{
    public static class ConfigManager
    {
        public static string Master_File { get; private set; } = string.Empty;
        public static string Template_File_Path { get; private set; } = string.Empty;
        public static string Template_File_Password { get; private set; } = string.Empty;
        public static string Output_Location { get; private set; } = string.Empty;
        public static string Unify_Folder { get; private set; } = string.Empty;
        public static string Done_Folder_Format { get; private set; } = "yyyyMMddHHmm_Done";
        public static string Report_File_Format { get; private set; } = "yyyyMMddHHmm_UnifyRpt.xlsx";
        public static int Workload_Main_Sheet { get; private set; } = 2;
        public static string Coursework_Text { get; private set; } = "N";
        public static string Research_Text { get; private set; } = "N";
        public static string Internship_Text { get; private set; } = "N";
        public static string DualCampus_Text { get; private set; } = "N";
        
        public static int Master_First_Data_Row { get; private set; } = 4;
        public static int Generate_From_Row { get; private set; } = 4;
        public static int Generate_To_Row { get; private set; } = 300;
        public static int Batch_Size { get; private set; } = 10;

        public static string SafesStaff_Label { get; private set; } = string.Empty;
        public static string TotalHrs_Label { get; private set; } = string.Empty;

        public static string OtherStaff_Label { get; private set; } = string.Empty;
        public static string Allocated_Overall_Address { get; private set; } = string.Empty;


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
                    config.TryGetValue("TemplateFilePath", out var templateFilePath);
                    config.TryGetValue("Output", out var output);
                    config.TryGetValue("Unify", out var unifyFolder);
                    config.TryGetValue("DoneFolderFormat", out var doneFolderFormat);
                    config.TryGetValue("ReportFileFormat", out var reportFileFormat);
                    config.TryGetValue("WorkloadMainSheet", out var workloadMainSheet);
                    config.TryGetValue("CourseworkText", out var courseworkText);
                    config.TryGetValue("ResearchText", out var researchText);
                    config.TryGetValue("InternshipText", out var internshipText);
                    config.TryGetValue("DualCampusText", out var dualcampusText);

                    config.TryGetValue("MasterDataFirstDataRow", out var masterFirstDataRow);
                    config.TryGetValue("GenerateFromRow", out var fromRow);
                    config.TryGetValue("GenerateToRow", out var toRow);
                    config.TryGetValue("BatchSizeToGenerate", out var batchSize);

                    config.TryGetValue("SafesStaffLabel", out var safesStaffLabel);
                    config.TryGetValue("OtherStaffLabel", out var otherStaffLabel);
                    config.TryGetValue("TotalHrsLabel", out var totalHrsLabel);
                    config.TryGetValue("AllocatedOverallAddress", out var allocatedOverallAddress);

                    config.TryGetValue("TemplateFilePassword", out var templateFilePassword);

                    Master_File = !string.IsNullOrEmpty(masterFile)
                        ? Path.Combine(AppContext.BaseDirectory, masterFile)
                        : string.Empty;

                    Template_File_Path = !string.IsNullOrEmpty(templateFilePath)
                        ? Path.Combine(AppContext.BaseDirectory, templateFilePath)
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
                    Research_Text = !string.IsNullOrEmpty(researchText) ? researchText : string.Empty;
                    Internship_Text = !string.IsNullOrEmpty(internshipText) ? internshipText : string.Empty;
                    DualCampus_Text = !string.IsNullOrEmpty(dualcampusText) ? dualcampusText : string.Empty;

                    Master_First_Data_Row = !string.IsNullOrEmpty(masterFirstDataRow) ?
                        System.Convert.ToInt32(masterFirstDataRow) : 0;
                    Generate_From_Row = !string.IsNullOrEmpty(fromRow) ?
                        System.Convert.ToInt32(fromRow) : Master_First_Data_Row;
                    Generate_To_Row = !string.IsNullOrEmpty(toRow) ?
                        System.Convert.ToInt32(toRow) : 300;

                    if (Generate_To_Row <= Generate_From_Row)
                        Generate_To_Row = Generate_From_Row + 1;

                    if (Generate_From_Row == 0) Generate_From_Row = Master_First_Data_Row;
                    if (Generate_To_Row == 0) Generate_To_Row = 300;

                    Batch_Size = !string.IsNullOrEmpty(batchSize) ?
                        System.Convert.ToInt32(batchSize) : 10;

                    SafesStaff_Label = !string.IsNullOrEmpty(safesStaffLabel) ?
                        safesStaffLabel : string.Empty;

                    OtherStaff_Label = !string.IsNullOrEmpty(otherStaffLabel) ?
                        otherStaffLabel : string.Empty;

                    TotalHrs_Label = !string.IsNullOrEmpty(totalHrsLabel) ?
                        totalHrsLabel : string.Empty;

                    Allocated_Overall_Address = !string.IsNullOrEmpty(allocatedOverallAddress) ?
                        allocatedOverallAddress : string.Empty;

                    Template_File_Password = !string.IsNullOrEmpty(templateFilePassword) ? templateFilePassword : string.Empty;
                }
                else
                {
                    Master_File = string.Empty;
                    Template_File_Path = string.Empty;
                    Output_Location = string.Empty;
                    Unify_Folder = string.Empty;
                    Done_Folder_Format = "yyyyMMddHHmm_Done";
                    Report_File_Format = "yyyyMMddHHmm_UnifyRpt.xlsx";

                    Coursework_Text = string.Empty;
                    Research_Text = string.Empty;
                    Internship_Text = string.Empty;
                    DualCampus_Text = string.Empty;

                    Master_First_Data_Row = 4;
                    Generate_From_Row = Master_First_Data_Row;
                    Generate_To_Row = 300;
                    Batch_Size = 10;

                    SafesStaff_Label = string.Empty;
                    OtherStaff_Label = string.Empty;
                    TotalHrs_Label = string.Empty;
                    Allocated_Overall_Address = string.Empty;

                    Template_File_Password = string.Empty;
                }
            }
            catch
            {
                Master_File = string.Empty;
                Template_File_Path = string.Empty;
                Output_Location = string.Empty;
                Unify_Folder = string.Empty;
                Done_Folder_Format = "yyyyMMddHHmm_Done";
                Report_File_Format = "yyyyMMddHHmm_UnifyRpt.xlsx";

                Coursework_Text = string.Empty;
                Research_Text = string.Empty;
                Internship_Text = string.Empty;
                DualCampus_Text = string.Empty;

                Master_First_Data_Row = 4;
                Generate_From_Row = Master_First_Data_Row;
                Generate_To_Row = 300;
                Batch_Size = 10;

                SafesStaff_Label = string.Empty;
                OtherStaff_Label = string.Empty;
                TotalHrs_Label = string.Empty;
                Allocated_Overall_Address = string.Empty;

                Template_File_Password = string.Empty;
            }
        }

        public static bool IsCourseWork(string sType)
        {
            return sType.ToLower().Trim().Contains(Coursework_Text.ToLower().Trim());
        }

        public static bool IsResearchInternship(string sType)
        {
            return sType.ToLower().Trim().Contains(Research_Text.ToLower().Trim()) ||
                   sType.ToLower().Trim().Contains(Internship_Text.ToLower().Trim());
        }

        public static bool IsDualCampus(string sType)
        {
            return sType.ToLower().Trim().Contains(DualCampus_Text.ToLower().Trim());
        }

        public static string GetTemplateFile(string sType)
        {
            if (IsCourseWork(sType))
                return Path.Combine(Template_File_Path, "standard-template.xlsx");
            else if (IsResearchInternship(sType))
                return Path.Combine(Template_File_Path, "research-template.xlsx");
            else if (IsDualCampus(sType))
                return Path.Combine(Template_File_Path, "dual-template.xlsx");
            else
                return string.Empty;
        }

    }
}
