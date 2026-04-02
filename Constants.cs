namespace WpfMailMerge;

public static class Constants
    {
    public const string VAR_RECIPIENTS = "Dko3Recepient";
    public const string VAR_ATTACHMENTS = "Dko3Attachments";
    public const string VAR_SUBJECT = "Dko3Subject";
    public const string APP_NAME = "MailMerge";
    public const string MERGED_FILE_PREFIX = "Mail";
    public const string WAITING = "Waiting...";
    public static string SENT_FILE_PREFIX = "sent_";
    public const string IDX_MARKER = "__IDX ";
    public const string INSERT_MARKER = "INSERT";
    public const string ATTACH_MARKER = "ATTACH";
    public const string SUBJECT_MARKER = "SUBJECT";
    public const string MAILTO_MARKER = "MAILTO";
    public const string COLLAPSE_MARKER = "COLLAPSE";
    public const string END_COLLAPSE_MARKER = "ENDCOLLAPSE";
    public const char SUBFIELD_SEPARATOR = '#';
    public const string MARKER_OPTION_NEWLINE = "+LINE";

    public static readonly string[] AllMarkerOptions = [MARKER_OPTION_NEWLINE];
    }

