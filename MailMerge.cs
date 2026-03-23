using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Channels;
using System.Windows;
using WpfMailMerge.Progress;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace WpfMailMerge;

internal partial class MailMerge
    {
    public event EventHandler? RunningStateChanged;
    public event EventHandler? RequestedStartIndexChanged;
    public event EventHandler? HasRecoveredStartIndexChanged;
    public event EventHandler? NamedRangesChanged;

    private Outlook.Application? outlook;

    private Dictionary<string, Word.Document> cachedWordDocs = new Dictionary<string, Word.Document>();
    const int batchLen = 20;
    private IProgressObservable progressListener;
    private CancellationTokenSource cancelToken;
    private Progress<Status> progressIndicator;
    private Channel<ExcelRequest> excelChannel;
    private Channel<MailFileInfo> mailDocChannel;

    public bool IsRunning { get; private set; } = false;
    public int RequestedStartIndex { get; internal set; } = 0;
    public bool HasRecoveredStartIndex { get; private set; } = false;
    public List<RangeDef> NamedRanges { get; private set; } = [new RangeDef {BookName="", Name="..", Range="",  RangeType = RangeType.Waiting }];

    public JsonSettings settings;

    public MailMerge()
        {
        this.progressListener = new DummyProgressObservable();
        this.settings = JsonSettings.Load();
        this.cancelToken = new CancellationTokenSource();
        this.progressIndicator = new Progress<Progress.Status>((status) => this.HandleThreadsProgress(status));
        this.excelChannel = Channel.CreateUnbounded<ExcelRequest>();
        this.mailDocChannel = Channel.CreateUnbounded<MailFileInfo>();

        this.StartThreads();
        }

    public void SetProgressObservable(IProgressObservable progressListener) => this.progressListener = progressListener;

    public void StartThreads()
        {
        var _ = Task.Run(() => this.HandleExcelRequests(this.settings, this.cancelToken.Token, this.progressIndicator, excelChannel.Reader));
        }

    private Outlook.Application GetOutlook()
        {
        if (this.outlook is null)
            this.outlook = new Outlook.Application();
        return this.outlook;
        }

    public List<MailAccount> GetSenders()
        {
        if(this.outlook is null)
            outlook = new Outlook.Application();
        List<MailAccount> accounts = new List<MailAccount>();
        for (int i = 1; i <= outlook.Session.Accounts.Count; i++)
            {
            accounts.Add(new MailAccount { DisplayName = outlook.Session.Accounts[i].DisplayName, Index = i });
            }
        return accounts;
        }

    public static List<MailAccount> GetSendersAsync()
        {
        var outlook = new Outlook.Application();
        List<MailAccount> accounts = new List<MailAccount>();
        for (int i = 1; i <= outlook.Session.Accounts.Count; i++)
            {
            accounts.Add(new MailAccount { DisplayName = outlook.Session.Accounts[i].DisplayName, Index = i });
            }
        Marshal.ReleaseComObject(outlook);
        outlook = null;
        GC.Collect();
        GC.WaitForPendingFinalizers();
        return accounts;
        }

    public void CheckRecovery()
        {
        if (!JsonRecovery.Exists())
            return;

        var jsonRecovery = JsonRecovery.Load();
        DateTime templateModifiedTime = File.GetLastWriteTime(this.settings.WordTemplateFileName);
        DateTime templateRecoveryTime = DateTime.ParseExact(jsonRecovery.TemplateDate, "O", CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.RoundtripKind);
        if (templateModifiedTime != templateRecoveryTime)
            {
            var answer = MessageBox.Show("Last task can be continued, but the template document has changed since then. Do you still want to continue the previous task?", "Mail merge recovery", MessageBoxButton.YesNo);
            if (answer == MessageBoxResult.Yes)
                SetRecoveredStartIndex();
            return;
            }
        DateTime dataModifiedTime = File.GetLastWriteTime(this.settings.DataSourceFileName);
        DateTime dataRecoveryTime = DateTime.ParseExact(jsonRecovery.DataDate, "O", CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.RoundtripKind);
        if (dataModifiedTime != dataRecoveryTime)
            {
            var answer = MessageBox.Show("Last task can be continued, but the data file has changed since then. Do you still want to continue the previous task?", "Mail merge recovery", MessageBoxButton.YesNo);
            if (answer == MessageBoxResult.Yes)
                SetRecoveredStartIndex();
            return;
            }
        SetRecoveredStartIndex();
        }

    private bool SetRecoveredStartIndex()
        {
        var files = Directory.GetFiles(DocBuilder.MergedDocsDir, Constants.SENT_FILE_PREFIX + "*.*");
        Array.Sort(files);
        var lastFile = files.Last();
        if (lastFile is null)
            {
            MessageBox.Show("Could not determine last file. Setting start position to 1.");
            return false;
            }
        var index = RxFirstInt().Match(lastFile).Value;
        if (index is null)
            {
            MessageBox.Show("Could not determine last file. Setting start position to 1.");
            return false;
            }
        this.SetStartIndex(int.Parse(index) + 1 + 1, true); //add 1 to start with the next file, and 1 to convert to 1-based.
        return true;
        }

    public void Start(JsonSettings settings)
        {
        SetRunningState(true);
        if (!this.PerformChecks(settings))
        {
            SetRunningState(false);
            return;
        }
        if (settings.NamedRange == null)
        {
            SetRunningState(false);
            return;
        }
        this.progressListener.ReportInfo("Preparing data...");
        this.excelChannel.Writer.TryWrite(new ExcelRequest(new DataParams(settings.DataSourceFileName, settings.NamedRange)));
        }

    private void SetRunningState(bool runningState)
    {
        this.IsRunning = runningState;
        this.RunningStateChanged?.Invoke(this, EventArgs.Empty);
    }
    private void SetStartIndex(int startIndex, bool recovered)
        {
        this.RequestedStartIndex = startIndex;
        this.RequestedStartIndexChanged?.Invoke(this, EventArgs.Empty);
        this.HasRecoveredStartIndex = recovered;
        this.HasRecoveredStartIndexChanged?.Invoke(this, EventArgs.Empty);
        }

    public void Stop()
        {
        this.cancelToken?.Cancel();
        }

    public void HandleThreadsProgress(Progress.Status status)
        {
        switch(status.StatusType)
            {
            case Progress.StatusType.Message:
                this.progressListener.ReportInfo(status.GetMessage());
                break;
            case Progress.StatusType.Progress:
                var progressInfo = status.GetProgressInfo();
                this.progressListener.ReportProgress(progressInfo.CurrentValue, progressInfo.MaxValue);
                break;
            case Progress.StatusType.Error:
                var error = status.GetError();
                this.progressListener.ReportError(error.message);
                break;
            case Progress.StatusType.ExcelRanges:                ;
                var ranges = status.GetExcelNamedRanges();
                this.NamedRanges = ranges;
                this.NamedRangesChanged?.Invoke(this, EventArgs.Empty);
                break;
            case Progress.StatusType.ExcelData:
                var data = status.GetExcelData();
                //assuming we pressed Start.
                data.Truncate(20); //todo: TEST!

                //IWordToEmailStrategy wordToEmail = new WordCopyPaste(settings);
                IWordToEmailStrategy wordToEmail = new WordToRtfEmail(settings);

                var _ = Task.Run(() => this.SendMails(settings, progressIndicator, mailDocChannel.Reader, wordToEmail));
                var __ = Task.Run(() => this.BuildTheDocs(settings, data, this.cancelToken.Token, progressIndicator, mailDocChannel.Writer, wordToEmail, this.RequestedStartIndex));
                SetRunningState(false);
                break;
            }
        }

    private void BuildTheDocs(JsonSettings settings, ExcelData excelData, CancellationToken cancelToken, IProgress<Progress.Status> progress, ChannelWriter<MailFileInfo> channelWriter, IWordToEmailStrategy wordToEmail, int startIndex)
        {
        progress.Report(new Progress.Status("Checking template file..."));
        var docBuilder = new DocBuilder(settings.WordTemplateFileName, excelData, wordToEmail);
        if(CancelBuild()) return;
        var checkResults = docBuilder.GetChecksResults();
        if (checkResults.Count > 0)
            {
            progress.Report(new Progress.Status(String.Join("\n", checkResults), -1));
            docBuilder.CloseAll();
            return;
            }
        progress.Report(new Progress.Status("Creating mail documents..."));

        CreateRecoveryFile(settings);

        for (int i = startIndex; i < excelData.Rows.Count; i++)
            {
            string fileName = docBuilder.BuildDoc(i);
            if (!channelWriter.TryWrite(new MailFileInfo { FileName = fileName, Index = i, Count = excelData.Rows.Count }))
                {
                throw new Exception("Can't write to channel.");
                }
            if (CancelBuild()) 
                return;
            }
        JsonRecovery.Delete();
        docBuilder.CloseAll();

        bool CancelBuild()
            {
            if (cancelToken.IsCancellationRequested)
                {
                this.progressListener.ReportInfo("Stopped...");
                docBuilder.CloseAll();
                }
            return cancelToken.IsCancellationRequested;
            }
        }

    private void CreateRecoveryFile(JsonSettings settings)
        {
        var templateDate = File.GetLastWriteTime(settings.WordTemplateFileName).ToString("O");
        var dataDate = File.GetLastWriteTime(settings.DataSourceFileName).ToString("O");
        JsonRecovery recovery = new() { TemplateDate = templateDate, DataDate = dataDate };
        recovery.Save();
        }

    private async void SendMails(JsonSettings settings, IProgress<Progress.Status> progress, ChannelReader<MailFileInfo> channelReader, IWordToEmailStrategy wordToEmail)
        {
        MailSender mailSender = new(settings, wordToEmail);
        mailSender.SetProgressObservable(progressListener);
        Debug.WriteLine("Waiting for mails to send...");
        try
            {
            while (true)
                {
                MailFileInfo mailFileInfo = await channelReader.ReadAsync();
                mailSender.SendOneMail(mailFileInfo.FileName);
                progress.Report(new Progress.Status(0, mailFileInfo.Count, mailFileInfo.Index + 1));
                if(mailFileInfo.Count == mailFileInfo.Index + 1)
                    progress.Report(new Progress.Status("All mails have been sent."));
                }
            }
        catch (ChannelClosedException e)
            {
            Debug.WriteLine("Mail channel closed.");
            mailSender.CloseAll();
            }
        }

    private async void HandleExcelRequests(JsonSettings settings, CancellationToken cancelToken, IProgress<Progress.Status> progress, ChannelReader<ExcelRequest> channelReader)
        {
        ExcelDataSource excelDataSource = new();
        excelDataSource.SetProgressObservable(progressListener);
        Debug.WriteLine("Waiting for excel requests...");
        try
            {
            while (true)
                {
                ExcelRequest request = await channelReader.ReadAsync();
                switch (request.requestType)
                    {
                    case ExcelRequestType.NamedRanges:
                        var namedRangesParams = request.GetNamedRangesParams();
                        var ranges = excelDataSource.GetRanges(namedRangesParams.filePath);
                        progress.Report(new Progress.Status(ranges));
                        break;
                    case ExcelRequestType.Data:
                        var dataParams = request.GetDataParams();
                        var data = excelDataSource.GetData(dataParams.filePath, dataParams.rangeName);
                        progress.Report(new Progress.Status(data));
                        break;
                    }

                if (cancelToken.IsCancellationRequested)
                    {
                    excelDataSource.CloseAll();
                    break;
                    }
                }
            }        
        catch (ChannelClosedException e)
            {
            Debug.WriteLine("Excel channel closed.");
            excelDataSource.CloseAll();
            }
        Debug.WriteLine("Excel thread ended.");
        }

    private bool PerformChecks(JsonSettings settings)
        {
        if (!File.Exists(settings.WordTemplateFileName))
            {
            MessageBox.Show($"Word template file not found: {settings.WordTemplateFileName}");
            return false;
            }
        if (!File.Exists(settings.DataSourceFileName))
            {
            MessageBox.Show($"Data source file not found: {settings.DataSourceFileName}");
            return false;
            }
        //check if merged file folder is empty
        if (!DocBuilder.IsMergeDirEmpty())
            {
            var reply = MessageBox.Show("Merged documents directory is not empty. Clear directory?", "Mail Merge", MessageBoxButton.YesNoCancel);
            if (reply == MessageBoxResult.Cancel)
                return false;
            if (reply == MessageBoxResult.Yes)
                {
                if(!DocBuilder.ClearMergeDir())
                    return false;
                }
            }
        return true;
        }

    public void CloseAll()
        {
        this.excelChannel.Writer.Complete();
        this.mailDocChannel.Writer.Complete();
        }

    ~MailMerge()
        {
        CloseAll();
        }

    [GeneratedRegex(@"\d+")]
    private static partial Regex RxFirstInt();

    internal void SetDataSourceFileName(string dataSourceFileName)
        {
        if(File.Exists(dataSourceFileName))
            {
            ExcelRequest req = new ExcelRequest(new RangesParams(dataSourceFileName));
            this.excelChannel.Writer.TryWrite(req);
            }
        }
    }

internal class DummyProgressObservable : IProgressObservable
    {
    public void ReportProgress(int value, int maxValue){}
    public void SetProgress(int value){}
    public void ReportError(string error){}
    public void ReportInfo(string info){}
    }

internal class MailFileInfo
    {
    public required string FileName;
    public required int Index;
    public required int Count;
    }


internal enum ExcelRequestType { NamedRanges, Data }

internal record RangesParams(string filePath);
internal record DataParams(string filePath, string rangeName);

internal class ExcelRequest
    {
    public readonly ExcelRequestType requestType;
    private readonly object Data;

    public ExcelRequest(RangesParams rangesParams) { this.requestType = ExcelRequestType.NamedRanges; this.Data = rangesParams;}
    public ExcelRequest(DataParams dataParams) { this.requestType = ExcelRequestType.Data; this.Data = dataParams;  }
    
    public RangesParams GetNamedRangesParams() { return (RangesParams) this.Data; }
    public DataParams GetDataParams() { return (DataParams) this.Data; }
    }