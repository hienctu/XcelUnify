using System.Text.Json;


namespace XcelUnify.Helpers
{
    public static class ConfigManager
    {
        public static string Master_File { get; private set; } = string.Empty;
        public static string Template_File { get; private set; } = string.Empty;
        public static string Output_Location { get; private set; } = string.Empty;
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

                    Master_File = !string.IsNullOrEmpty(masterFile)
                        ? Path.Combine(AppContext.BaseDirectory, masterFile)
                        : string.Empty;

                    Template_File = !string.IsNullOrEmpty(templateFile)
                        ? Path.Combine(AppContext.BaseDirectory, templateFile)
                        : string.Empty;

                    Output_Location = !string.IsNullOrEmpty(output)
                        ? Path.Combine(AppContext.BaseDirectory, output)
                        : string.Empty;
                }
                else
                {
                    Master_File = string.Empty;
                    Template_File = string.Empty;
                    Output_Location = string.Empty;
                }
            }
            catch
            {
                Master_File = string.Empty;
                Template_File = string.Empty;
                Output_Location = string.Empty;
            }
        }

    }
}
