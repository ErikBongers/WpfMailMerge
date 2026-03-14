using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Threading;

namespace WpfMailMerge;

public class MailAccount
    {
    public required string DisplayName { get; set; }
    public required int Index { get; set; }
    }

public interface IProgressObservable
    {
    void ReportProgress(int value, int maxValue, string info);
    void SetProgress(int value);
    }

public class MailMergeViewModel : INotifyPropertyChanged, IProgressObservable
    {
    #region Properties for data binding
    public event PropertyChangedEventHandler? PropertyChanged;
    void OnPropertyChanged(string propName)
        {
        if (this.PropertyChanged != null)
            this.PropertyChanged(this, new PropertyChangedEventArgs(propName));
        }

    private string statusMessage = "<Status message>";
    public string StatusMessage
        {
        get { return statusMessage; }
        set { statusMessage = value; OnPropertyChanged(nameof(StatusMessage)); }
        }
    private int progressValue = 0;
    public int ProgressValue
        {
        get { return progressValue; }
        set { progressValue = value; OnPropertyChanged(nameof(ProgressValue)); }
        }
    private int progressMaxValue = 100;
    public int ProgressMaxValue
        {
        get { return progressMaxValue; }
        set { progressMaxValue = value; OnPropertyChanged(nameof(ProgressMaxValue)); }
        }
    private string progressInfo = "<Progress info>";
    public string ProgressInfo
        {
        get { return progressInfo; }
        set { progressInfo = value; OnPropertyChanged(nameof(ProgressInfo)); }
        }
    public NotifyTaskCompletion<List<MailAccount>> mailAccounts;
    public NotifyTaskCompletion<List<MailAccount>> MailAccounts
        {
        get { return mailAccounts; }
        set
            {
            mailAccounts = value;
            OnPropertyChanged(nameof(MailAccounts));
            OnPropertyChanged(nameof(SenderComboVisibility));
            }
        }
    private int mailAccountIndex = -1;
    public int MailAccountIndex
        {
        get { return mailAccountIndex; }
        set { mailAccountIndex = value; OnPropertyChanged(nameof(MailAccountIndex)); }
        }
    private bool useTestRecipient = true;
    public bool UseTestRecipient
        {
        get { return useTestRecipient; }
        set { 
            useTestRecipient = value; OnPropertyChanged(nameof(UseTestRecipient)); 
            TestEmailVisibility = value ? Visibility.Visible : Visibility.Hidden;
            }
        }
    private string testRecipient = "erikbongers@outlook.com";
    public string TestRecipient
        {
        get { return testRecipient; }
        set { testRecipient = value; OnPropertyChanged(nameof(TestRecipient)); }
        }

    private int savedMailAccountIndex;
    private Visibility testEmailVisibility = Visibility.Visible;
    public Visibility TestEmailVisibility
        {
        get { return testEmailVisibility; }
        set { testEmailVisibility = value; OnPropertyChanged(nameof(TestEmailVisibility)); }
        }
    private string wordTemplateFileName = @"C:\Users\erikb\Desktop\MailMerge.docm";
    public string WordTemplateFileName
        {
        get { return wordTemplateFileName; }
        set { wordTemplateFileName = value; OnPropertyChanged(nameof(WordTemplateFileName)); }
        }
    private string dataSourceFileName = @"C:\Users\erikb\Desktop\TestDataMailMergeV2.xlsm";
    public string DataSourceFileName
        {
        get { return dataSourceFileName; }
        set { dataSourceFileName = value; OnPropertyChanged(nameof(DataSourceFileName)); }
        }
    private string onBehalfOfEmail = "academie.berchem.muziek.woord@stedelijkonderwijs.be";
    public string OnBehalfOfEmail 
        {
        get { return onBehalfOfEmail; }
        set { onBehalfOfEmail = value; OnPropertyChanged(nameof(OnBehalfOfEmail)); }
        }
    public Visibility SenderComboVisibility {
        get { return this.MailAccounts?.Result?.Count == 1 ? Visibility.Collapsed : Visibility.Visible; }
        }
    #endregion

    private const string SETTINGS_FILENAME = "settings.json";

    private readonly MailMerge mailMerge = new();

    public MailMergeViewModel()
        {
        mailMerge.SetProgressObservable(this);
        LoadJsonSettings();
        mailAccounts = new NotifyTaskCompletion<List<MailAccount>>(LoadMailAccounts(), new List<MailAccount>([new MailAccount { DisplayName = "Loading...", Index = -1 }]));
        mailAccounts.PropertyChanged += (s, e) =>
            {
            if (e.PropertyName == nameof(MailAccounts.IsCompleted))
                {
                OnPropertyChanged(nameof(SenderComboVisibility));
                    this.MailAccountIndex = this.savedMailAccountIndex;
                    Task.Delay(10).ContinueWith(t =>
                    {
                        this.MailAccountIndex = this.savedMailAccountIndex;
                    });
                }
            };
        }

    public static async Task<List<MailAccount>> LoadMailAccounts()
        {
        return await Task.Run(() => MailMerge.GetSendersAsync());

        //todo
        //this.MailAccounts = this.mailMerge.GetSenders();
        //if (this.mailAccounts.Count == 1)
        //    {
        //    this.MailAccountIndex = this.mailAccounts[0].Index;
        //    }
        }

    private void LoadJsonSettings()
        {
        string localDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        string settingsFile = Path.Combine(localDir, MailMerge.APP_NAME, SETTINGS_FILENAME);
        if (File.Exists(settingsFile))
            {
            string json = File.ReadAllText(settingsFile);
            // Deserialize json to load settings
            JsonSettings? settings = System.Text.Json.JsonSerializer.Deserialize<JsonSettings>(json);
            if (settings == null)
                return;
            this.WordTemplateFileName = settings.WordTemplateFileName;
            this.DataSourceFileName = settings.DataSourceFileName;
            this.UseTestRecipient = settings.UseTestRecipient;
            this.TestRecipient = settings.TestRecipient;
            this.savedMailAccountIndex = settings.MailAccountIndex??-1;
            this.MailAccountIndex = -1; //until we can load the mail accounts.
            }
        }

    public void SaveJsonSettings()
        {
        string localDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        string appDir = Path.Combine(localDir, MailMerge.APP_NAME);
        if (!Directory.Exists(appDir))
            Directory.CreateDirectory(appDir);
        string settingsFile = Path.Combine(appDir, SETTINGS_FILENAME);
        JsonSettings settings = ScrapeSettings();
        string json = System.Text.Json.JsonSerializer.Serialize(settings);
        File.WriteAllText(settingsFile, json);
        }

    private JsonSettings ScrapeSettings()
        {
        return new JsonSettings
            {
            WordTemplateFileName = this.WordTemplateFileName,
            DataSourceFileName = this.DataSourceFileName,
            UseTestRecipient = this.UseTestRecipient,
            TestRecipient = this.TestRecipient,
            MailAccountIndex = this.MailAccountIndex,
            OnBehalfOfEmail = this.OnBehalfOfEmail
            };
        }

    public void Start()
        {
        this.MailAccountIndex = 2;
        //this.mailMerge.Start(ScrapeSettings());
        }

    public void ReportProgress(int value, int maxValue, string info)
        {
        this.ProgressMaxValue = maxValue;
        this.ProgressValue = value;
        this.StatusMessage = info;
        int percentage = (int)((double)ProgressValue / ProgressMaxValue * 100);
        this.ProgressInfo = $"{ProgressValue} ({percentage}%) of {ProgressMaxValue}";
        }

    public void SetProgress(int value)
        {
        this.ProgressValue = value;
        }
    }
