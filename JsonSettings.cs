namespace WpfMailMerge;

public class JsonSettings
    {
    public required string WordTemplateFileName { get; set; }
    public required string DataSourceFileName { get; set; }
    public required bool UseTestRecipient { get; set; }
    public required string TestRecipient { get; set; }
    public int? MailAccountIndex { get; set; }
    public string? OnBehalfOfEmail { get; set; }
    }
