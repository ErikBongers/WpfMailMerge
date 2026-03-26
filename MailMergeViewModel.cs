using System.ComponentModel;
using System.IO;
using System.Windows;

namespace WpfMailMerge;

public class MailAccount
    {
    public required string DisplayName { get; set; }
    public required int Index { get; set; }
    }

public interface IProgressObservable
    {
    void ReportProgress(int value, int maxValue);
    void SetProgress(int value);
    void ReportError(string error);
    void ReportInfo(string info);
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
    public Visibility SenderComboVisibility
        {
        get { return this.MailAccounts?.Result?.Count == 1 ? Visibility.Collapsed : Visibility.Visible; }
        }
    private int savedMailAccountIndex;

    public List<RangeDef> namedRanges = [new RangeDef { BookName= "", SheetName="", Name = "..", Range = "", RangeType = RangeType.Waiting }];
    public List<RangeDef> NamedRanges
        {
        get { return namedRanges; }
        set
            {
            namedRanges = value;
            OnPropertyChanged(nameof(NamedRanges));
            OnPropertyChanged(nameof(NamedRangesComboVisibility));
            }
        }
    private string selectedNamedRange = Constants.WAITING;
    public string SelectedNamedRange
        {
        get { return selectedNamedRange; }
        set { 
            selectedNamedRange = value; 
            OnPropertyChanged(nameof(SelectedNamedRange));
            OnPropertyChanged(nameof(CanStart));
            }
        }

    public Visibility NamedRangesComboVisibility
        {
        get { return this.NamedRanges.Count == 1 ? Visibility.Collapsed : Visibility.Visible; }
        }
    private string savedNamedRange = Constants.WAITING;
    private bool useTestRecipient = true;
    public bool UseTestRecipient
        {
        get { return useTestRecipient; }
        set { 
            useTestRecipient = value; 
            OnPropertyChanged(nameof(UseTestRecipient)); 
            TestEmailVisibility = value ? Visibility.Visible : Visibility.Hidden;
            OnPropertyChanged(nameof(CanStart));
            }
        }
    private string testRecipient = "erikbongers@outlook.com";
    public string TestRecipient
        {
        get { return testRecipient; }
        set { 
            testRecipient = value; 
            OnPropertyChanged(nameof(TestRecipient));
            OnPropertyChanged(nameof(CanStart));
            }
        }

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
        set { 
            wordTemplateFileName = value; 
            OnPropertyChanged(nameof(WordTemplateFileName)); 
            OnPropertyChanged(nameof(CanStart));
            OnPropertyChanged(nameof(IsDocPathValid));
            }
        }
    private string dataSourceFileName = @"C:\Users\erikb\Desktop\TestDataMailMergeV2.xlsm";
    public string DataSourceFileName
        {
        get { return dataSourceFileName; }
        set { 
            dataSourceFileName = value; 
            OnPropertyChanged(nameof(DataSourceFileName));
            OnPropertyChanged(nameof(CanStart));
            OnPropertyChanged(nameof(IsDataPathValid));
            this.mailMerge.SetDataSourceFileName(this.dataSourceFileName);
            }
        }
    private string onBehalfOfEmail = "academie.berchem.muziek.woord@stedelijkonderwijs.be";
    public string OnBehalfOfEmail 
        {
        get { return onBehalfOfEmail; }
        set { onBehalfOfEmail = value; OnPropertyChanged(nameof(OnBehalfOfEmail)); }
        }

    public bool CanStart
        {
        get { 
            return this.mailAccountIndex >= 0 
                && !string.IsNullOrEmpty(this.wordTemplateFileName) 
                && !string.IsNullOrEmpty(this.dataSourceFileName) 
                && (this.UseTestRecipient == false || !string.IsNullOrEmpty(this.TestRecipient))
                && this.SelectedNamedRange != Constants.WAITING
                && this.IsDataPathValid
                && this.IsDocPathValid;
            }
        }
    private bool inError = false;
    public bool InError
        {
        get { return inError; }
        set { inError = value; OnPropertyChanged(nameof(InError)); }
        }
    private bool isRunning = false;
    public bool IsRunning
        {
        get { return isRunning; }
        set { 
            isRunning = value; 
            OnPropertyChanged(nameof(IsRunning)); 
            OnPropertyChanged(nameof(StartStopText)); 
            }
        }
    public string StartStopText
        {
        get { return isRunning ? "Stop" : "Start"; }
        }
    private int requestedStartIndex = 1; //1-based, because user indexing.
    public int RequestedStartIndex
        {
        get { return requestedStartIndex; }
        set { 
            requestedStartIndex = value; 
            OnPropertyChanged(nameof(RequestedStartIndex));
            this.mailMerge.RequestedStartIndex = this.requestedStartIndex-1; //convert to zero based.
            }
        }
    public Visibility showRecoveredStartIndexMessage = Visibility.Hidden;
    public Visibility ShowRecoveredStartIndexMessage
        {
        get { return showRecoveredStartIndexMessage; }
        set
            {
            showRecoveredStartIndexMessage = value;
            OnPropertyChanged(nameof(ShowRecoveredStartIndexMessage));
            }
        }
    private bool mergeOtherExcels = false;
    public bool MergeOtherExcels
        {
        get { return mergeOtherExcels; }
        set
            {
            mergeOtherExcels = value;
            OnPropertyChanged(nameof(MergeOtherExcels));
            }
        }
    public bool IsDocPathValid { get { return IsFileOfType(this.WordTemplateFileName,[".docx", ".docm"]); } }
    public bool IsDataPathValid { get { return IsFileOfType(this.DataSourceFileName, [".xlsx", ".xlsm"]); } }
    #endregion

    private bool IsFileOfType(string filePath, IEnumerable<string> extensions)
        {
        if (!File.Exists(filePath))
            return false;
        var fileExt = Path.GetExtension(filePath);
        
        return extensions.Contains<string>(fileExt);
        }

    private readonly MailMerge mailMerge;

    public MailMergeViewModel()
        {
        mailMerge = new(this);
        mailMerge.SetProgressObservable(this);
        mailMerge.RunningStateChanged += (_, _) => {
            this.IsRunning = this.mailMerge.IsRunning;
        };
        mailMerge.RequestedStartIndexChanged += (_, _) => {
            this.RequestedStartIndex = this.mailMerge.RequestedStartIndex;
        };
        mailMerge.HasRecoveredStartIndexChanged += (_, _) => {
            this.ShowRecoveredStartIndexMessage = this.mailMerge.HasRecoveredStartIndex ? Visibility.Visible : Visibility.Hidden;
        };
        mailMerge.NamedRangesChanged += (_, _) => {
            this.NamedRanges = this.mailMerge.NamedRanges;
            Task.Delay(10).ContinueWith(t =>
                {
                    this.SelectedNamedRange = this.savedNamedRange;
                });
        };
        LoadJsonSettings();
        mailAccounts = new NotifyTaskCompletion<List<MailAccount>>(LoadMailAccounts(), new List<MailAccount>([new MailAccount { DisplayName = "Loading...", Index = -1 }]));
        mailAccounts.PropertyChanged += (s, e) =>
        {
            if (e.PropertyName == nameof(MailAccounts.IsCompleted))
                {
                OnPropertyChanged(nameof(SenderComboVisibility));
                int newMailAccountIndex = -1;
                if (mailAccounts.Result?.Count == 1)
                    {
                    newMailAccountIndex = mailAccounts.Result[0].Index;
                    }
                else if (this.savedMailAccountIndex < mailAccounts.Result?.Count)
                    {
                    newMailAccountIndex = this.savedMailAccountIndex;
                    }
                Task.Delay(10).ContinueWith(t =>
                {
                    this.MailAccountIndex = newMailAccountIndex;
                });
                }
        };
        }

    private static Task<List<MailAccount>> LoadMailAccounts() //todo: can I remove the await and just return the tast from Task.Run()?
        {
        return Task.Run(() => MailMerge.GetSendersAsync());
        }

    public void LoadJsonSettings()
        {
        var settings = this.mailMerge.LoadSettings();
        this.WordTemplateFileName = settings.WordTemplateFileName;
        this.DataSourceFileName = settings.DataSourceFileName;
        this.UseTestRecipient = settings.UseTestRecipient;
        this.TestRecipient = settings.TestRecipient;
        this.savedMailAccountIndex = settings.MailAccountIndex;
        this.MailAccountIndex = -1; //until we can load the mail accounts.
        this.savedNamedRange = settings.NamedRange ?? "";
        this.MergeOtherExcels = settings.MergeOtherExcels;
        }


    public void SaveJsonSettings()
        {
        ScrapeSettings().Save();
        }

    public JsonSettings ScrapeSettings()
        {
        return new JsonSettings
            {
            WordTemplateFileName = this.WordTemplateFileName,
            DataSourceFileName = this.DataSourceFileName,
            UseTestRecipient = this.UseTestRecipient,
            TestRecipient = this.TestRecipient,
            MailAccountIndex = this.MailAccountIndex,
            OnBehalfOfEmail = this.OnBehalfOfEmail,
            NamedRange = this.SelectedNamedRange,
            MergeOtherExcels = this.MergeOtherExcels
            };
        }

    public void StartStop()
        {
        this.SaveJsonSettings();
        if (this.mailMerge.IsRunning)
            this.mailMerge.Stop();
        else
            this.mailMerge.Start(ScrapeSettings());
        }

    public void ReportProgress(int value, int maxValue)
        {
        this.ProgressMaxValue = maxValue;
        this.ProgressValue = value;
        int percentage = (int)((double)ProgressValue / ProgressMaxValue * 100);
        this.ProgressInfo = $"{ProgressValue} ({percentage}%) of {ProgressMaxValue}";
        }

    public void SetProgress(int value)
        {
        this.ProgressValue = value;
        }

    public void ReportError(string error)
        {
        this.StatusMessage = error;
        this.InError = true;
        }

    public void ReportInfo(string info)
        {
        this.StatusMessage = info;
        }

    public void CloseAll()
        {
        this.mailMerge.CloseAll();
        }

    internal void CheckRecovery()
        {
        this.mailMerge.CheckRecovery();
        }

    ~MailMergeViewModel()
        {
        this.CloseAll();
        }
    }
