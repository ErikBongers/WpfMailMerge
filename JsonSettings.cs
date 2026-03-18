using System.ComponentModel;
using System.IO;

namespace WpfMailMerge;

public class JsonSettings
    {
    public string WordTemplateFileName { get; set; } = "";
    public string DataSourceFileName { get; set; } = "";
    public bool UseTestRecipient { get; set; } = true;
    public string TestRecipient { get; set; } = "";
    public int MailAccountIndex { get; set; } = -1;
    public string OnBehalfOfEmail { get; set; } = "";
    public string NamedRange { get; set; } = "";
    public int DelayAfterClipboardCopy { get; set; } = 500;

    const string SETTINGS_FILENAME = "settings.json";

    public static JsonSettings LoadJsonSettings()
        {
        string localDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        string settingsFile = Path.Combine(localDir, Constants.APP_NAME, SETTINGS_FILENAME);
        string json = """{"WordTemplateFileName":"sdf"}""";
        if (File.Exists(settingsFile))
            json = File.ReadAllText(settingsFile);
        JsonSettings settings = System.Text.Json.JsonSerializer.Deserialize<JsonSettings>(json)!;
        return settings;
        }

    public void SaveJsonSettings()
        {
        string localDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        string appDir = Path.Combine(localDir, Constants.APP_NAME);
        if (!Directory.Exists(appDir))
            Directory.CreateDirectory(appDir);
        string settingsFile = Path.Combine(appDir, SETTINGS_FILENAME);
        string json = System.Text.Json.JsonSerializer.Serialize(this);
        File.WriteAllText(settingsFile, json);
        }
    }