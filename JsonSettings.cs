using System.ComponentModel;
using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace WpfMailMerge;

public static class JsonFile<T>
    {
    public static T Load(string baseFileName)
        {
        string appDir = GetFullAppDirPath();
        string fileName = Path.Combine(appDir, baseFileName);
        string json = """{}""";
        if (File.Exists(fileName))
            json = File.ReadAllText(fileName);
        return System.Text.Json.JsonSerializer.Deserialize<T>(json)!;
        }

    private static string GetFullAppDirPath()
        {
        string localDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        return Path.Combine(localDir, Constants.APP_NAME);
        }

    public static string CreateDirAndGetFileName(string fileName)
        {
        string appDir = GetFullAppDirPath();
        if (!Directory.Exists(appDir))
            Directory.CreateDirectory(appDir);
        return Path.Combine(appDir, fileName);
        }

    public static bool Exists(string fileName)
        {
        string appDir = GetFullAppDirPath();
        if (!Directory.Exists(appDir))
            return false;
        return File.Exists(Path.Combine(appDir, fileName));
        }

    public static void Save(T jsonObject, string baseFileName)
        {
        string settingsFile = JsonFile<JsonRecovery>.CreateDirAndGetFileName(baseFileName);
        var options = new JsonSerializerOptions { Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping };
        string json = System.Text.Json.JsonSerializer.Serialize(jsonObject, options);
        File.WriteAllText(settingsFile, json);
        }

    }

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
    public bool MergeOtherExcels { get; set; } = false;

    const string FILENAME = "settings.json";

    public static JsonSettings Load()
        {
        return JsonFile<JsonSettings>.Load(FILENAME);
        }

    public void Save()
        {
        JsonFile<JsonSettings>.Save(this, FILENAME);
        }
    }

public class JsonRecovery
    {
    public string TemplateDate { get; set; } = "";
    public string DataDate { get; set; } = "";

    const string FILENAME = "recovery.json";

    public static JsonRecovery Load()
        {
        return JsonFile<JsonRecovery>.Load(FILENAME);
        }

    public static bool Exists()
        {
        return JsonFile<JsonRecovery>.Exists(FILENAME);
        }

    public void Save()
        {
        JsonFile<JsonRecovery>.Save(this, FILENAME);
        }

    public static void Delete()
        {
        File.Delete(JsonFile<JsonRecovery>.CreateDirAndGetFileName(FILENAME));
        }
    }
