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

    public NotifyTaskCompletion<List<RangeDef>> namedRanges;
    public NotifyTaskCompletion<List<RangeDef>> NamedRanges
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
        get { return this.NamedRanges?.Result?.Count == 1 ? Visibility.Collapsed : Visibility.Visible; }
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
                && this.SelectedNamedRange != Constants.WAITING;
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
    #endregion

    private readonly MailMerge mailMerge = new();

    public MailMergeViewModel()
        {
        mailMerge.SetProgressObservable(this);
        mailMerge.RunningStateChanged += (_, _) => {
            this.IsRunning = this.mailMerge.IsRunning;
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
        namedRanges = new NotifyTaskCompletion<List<RangeDef>>(LoadDataSourceRanges(this.DataSourceFileName, this.mailMerge), new List<RangeDef>([new RangeDef{ Name = "..", Range="", RangeType=RangeType.Waiting }]));
        namedRanges.PropertyChanged += (s, e) =>
            {
            if (e.PropertyName == nameof(NamedRanges.IsCompleted))
                {
                OnPropertyChanged(nameof(NamedRangesComboVisibility));
                this.SelectedNamedRange = this.savedNamedRange;
                Task.Delay(10).ContinueWith(t =>
                    {
                    this.SelectedNamedRange= this.savedNamedRange;
                    });
                }
            };
        }

    private static Task<List<MailAccount>> LoadMailAccounts() //todo: can I remove the await and just return the tast from Task.Run()?
        {
        return Task.Run(() => MailMerge.GetSendersAsync());
        }

    private static Task<List<RangeDef>> LoadDataSourceRanges(string dataSourceFileName, MailMerge mailMerge)
        {
        return Task.Run(() =>
            {
                ExcelDataSource excelDataSource = mailMerge.SetExcelDataSource(dataSourceFileName);
                return excelDataSource.GetRanges();
                //return new List<RangeDef>([new RangeDef { Name = "..", Range = "", RangeType = RangeType.Waiting }, new RangeDef { Name = "sdfsd", Range = "", RangeType = RangeType.Waiting }]);
            });
        }

    public void LoadJsonSettings()
        {
        var settings = JsonSettings.Load();
        this.WordTemplateFileName = settings.WordTemplateFileName;
        this.DataSourceFileName = settings.DataSourceFileName;
        this.UseTestRecipient = settings.UseTestRecipient;
        this.TestRecipient = settings.TestRecipient;
        this.savedMailAccountIndex = settings.MailAccountIndex;
        this.MailAccountIndex = -1; //until we can load the mail accounts.
        this.savedNamedRange = settings.NamedRange ?? "";
        }


    public void SaveJsonSettings()
        {
        ScrapeSettings().Save();
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
            OnBehalfOfEmail = this.OnBehalfOfEmail,
            NamedRange = this.SelectedNamedRange
            };
        }

    public async Task StartStopAsync()
        {
        this.SaveJsonSettings();
        if (this.mailMerge.IsRunning)
            this.mailMerge.Stop();
        else
            await this.mailMerge.StartAsync(ScrapeSettings());
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

    ~MailMergeViewModel()
        {
        this.CloseAll();
        }
    }
