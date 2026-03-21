using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Channels;
using System.Windows;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace WpfMailMerge;

internal class MailMerge
    {
    public event EventHandler? RunningStateChanged;

    private Outlook.Application? outlook;

    private Dictionary<string, Word.Document> cachedWordDocs = new Dictionary<string, Word.Document>();
    const int batchLen = 20;
    private IProgressObservable progressListener;
    ExcelDataSource? excelDataSource;
    private CancellationTokenSource? cancelToken;

    public bool IsRunning { get; private set; } = false;
    public int RequestedStartIndex { get; internal set; } = 0;

    public MailMerge()
        {
        this.progressListener = new DummyProgressObservable();
        }

    public void SetProgressObservable(IProgressObservable progressListener) => this.progressListener = progressListener;

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

    public ExcelDataSource SetExcelDataSource(string rangeName)
        {
        this.excelDataSource = new ExcelDataSource(rangeName);
        return this.excelDataSource;
        }

    public async Task StartAsync(JsonSettings settings)
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
        if (this.excelDataSource is null)
        {
            SetRunningState(false);
            return;
        }
        if (this.excelDataSource is null)
        {
            SetRunningState(false);
            return;
        }
        this.progressListener.ReportInfo("Preparing data..."); //todo: put in await Task.Run()
        var data = this.excelDataSource.GetData(settings.NamedRange);
        data.Truncate(20); //todo: TEST!
        this.excelDataSource.CloseExcel();

        var channel = Channel.CreateUnbounded<string>();

        //IWordToEmailStrategy wordToEmail = new WordCopyPaste(settings);
        IWordToEmailStrategy wordToEmail = new WordToRtfEmail(settings);

        var progressIndicator = new Progress<Progress.Status>((status) => this.ReportDocsProgress(status));
        cancelToken = new CancellationTokenSource();
        var _ = Task.Run(() => this.SendMails(settings, cancelToken.Token, progressIndicator, channel.Reader, wordToEmail));
        await Task.Run(() => this.BuildTheDocs(settings, data, this.cancelToken.Token, progressIndicator, channel.Writer, wordToEmail, this.RequestedStartIndex));

        SetRunningState(false);
        }

    private void SetRunningState(bool runningState)
    {
        this.IsRunning = runningState;
        this.RunningStateChanged?.Invoke(this, EventArgs.Empty);
    }

    public void Stop()
        {
        this.cancelToken?.Cancel();
        }

    void ReportDocsProgress(Progress.Status status)
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
            }
        }

    private void BuildTheDocs(JsonSettings settings, ExcelData excelData, CancellationToken cancelToken, IProgress<Progress.Status> progress, ChannelWriter<string> channelWriter, IWordToEmailStrategy wordToEmail, int startIndex)
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

        for(int i = startIndex; i <excelData.Rows.Count; i++)
            {
            string fileName = docBuilder.BuildDoc(i);
            if (!channelWriter.TryWrite(fileName))
                {
                throw new Exception("Can't write to channel.");
                }
            progress.Report(new Progress.Status(0, excelData.Rows.Count, i+1));
            if (CancelBuild()) 
                return;
            }
        progress.Report(new Progress.Status("Finished creating documents."));
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

    private async void SendMails(JsonSettings settings, CancellationToken cancelToken, IProgress<Progress.Status> progress, ChannelReader<string> channelReader, IWordToEmailStrategy wordToEmail)
        {
        MailSender mailSender = new(settings, wordToEmail);
        mailSender.SetProgressObservable(progressListener);
        Debug.WriteLine("Waiting for mails to send...");
        while(true)
            {
            string fileName = await channelReader.ReadAsync();
            mailSender.SendOneMail(fileName);
            if (cancelToken.IsCancellationRequested)
                {
                mailSender.CloseAll();
                break;
                }
            }
        Debug.WriteLine("SendMails ended.");
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
        this.excelDataSource?.CloseExcel();
        }

    ~MailMerge()
        {
        CloseAll();
        }
    }

internal class DummyProgressObservable : IProgressObservable
    {
    public void ReportProgress(int value, int maxValue){}
    public void SetProgress(int value){}
    public void ReportError(string error){}
    public void ReportInfo(string info){}
    }
