using Microsoft.VisualBasic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Threading;
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

    public bool IsRunning {
        get;
        private set;
        }

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
        //outlook.Quit();
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
        this.IsRunning = true;
        this.RunningStateChanged?.Invoke(this, EventArgs.Empty);
        if (!this.PerformChecks(settings))
            return;
        if (settings.NamedRange == null)
            return;
        if (this.excelDataSource is null)
            return;
        if (this.excelDataSource is null)
            return;
        this.progressListener.ReportInfo("Preparing data...");
        AllowUIToUpdate(); //todo: put in task and get rid of this.
        var data = this.excelDataSource.GetData(settings.NamedRange);
        this.excelDataSource.CloseExcel();

        var progressIndicator = new Progress<DocServer.Status>((status) => this.ReportDocsProgress(status));
        cancelToken = new CancellationTokenSource();
        await Task.Run(() => this.BuildTheDocs(settings, data, this.cancelToken.Token, progressIndicator));
        this.IsRunning = false;
        this.RunningStateChanged?.Invoke(this, EventArgs.Empty);
        }

    public void Stop()
        {
        this.cancelToken?.Cancel();
        }

    void ReportDocsProgress(DocServer.Status status)
        {
        switch(status.StatusType)
            {
            case DocServer.StatusType.Message:
                this.progressListener.ReportInfo(status.GetMessage());
                break;
            case DocServer.StatusType.Progress:
                var progressInfo = status.GetProgressInfo();
                this.progressListener.ReportProgress(progressInfo.CurrentValue, progressInfo.MaxValue, "todo...");
                break;
            case DocServer.StatusType.Error:
                var error = status.GetError();
                this.progressListener.ReportError(error.message);
                break;
            }
        }

    private void BuildTheDocs(JsonSettings settings, ExcelData excelData, CancellationToken cancelToken, IProgress<DocServer.Status> progress)
        {
        progress.Report(new DocServer.Status("Checking template file..."));
        var docBuilder = new DocBuilder(settings.WordTemplateFileName, excelData);
        if(CancelBuild()) return;
        var checkResults = docBuilder.GetChecksResults();
        if (checkResults.Count > 0)
            {
            progress.Report(new DocServer.Status(String.Join("\n", checkResults), -1));
            docBuilder.CloseAll();
            return;
            }
        progress.Report(new DocServer.Status("Creating mail documents..."));

        for(int i = 1; i <=excelData.Rows.Count; i++)
            {
            progress.Report(new DocServer.Status(0, excelData.Rows.Count, i));
            docBuilder.BuildDoc(i);
            if (CancelBuild()) 
                return;
            }
        progress.Report(new DocServer.Status("Finished creating documents."));
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


    //Source - https://stackoverflow.com/a/73181682
    //Posted by Yaron Binder, modified by community.See post 'Timeline' for change history
    //Retrieved 2026-03-17, License - CC BY-SA 4.0
    private static void AllowUIToUpdate()
        {
        DispatcherFrame frame = new();
        // DispatcherPriority set to Input, the highest priority
        Dispatcher.CurrentDispatcher.Invoke(DispatcherPriority.Input, new DispatcherOperationCallback(delegate (object parameter)
            {
                frame.Continue = false;
                Thread.Sleep(20); // Stop all processes to make sure the UI update is perform
                return null;
                }), null);
        Dispatcher.PushFrame(frame);
        // DispatcherPriority set to Input, the highest priority
        Application.Current.Dispatcher.Invoke(DispatcherPriority.Input, new Action(delegate { }));
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
                DocBuilder.ClearMergeDir();
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
    public void ReportProgress(int value, int maxValue, string info){}
    public void SetProgress(int value){}
    public void ReportError(string error){}
    public void ReportInfo(string info){}
    }
