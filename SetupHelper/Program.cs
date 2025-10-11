using System;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "appSettings.json");

        if (!File.Exists(filePath))
        {
            Console.WriteLine("File not found: " + filePath);
            return;
        }

        string content = File.ReadAllText(filePath);
        string userProfile = Environment.GetEnvironmentVariable("USERPROFILE");
        if (string.IsNullOrEmpty(userProfile))
        {
            userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        }
        userProfile = userProfile.Replace(@"\", @"\\");
        content = content.Replace("<<your-profile-path>>", userProfile);
      
        content = content.Replace("<<your-profile-path>>", userProfile);
        File.WriteAllText(filePath, content);

        Console.WriteLine("Updated appSettings.json successfully.");
    }
}