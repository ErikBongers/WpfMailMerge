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
    private Outlook.Application? outlook;

    private Dictionary<string, Word.Document> cachedWordDocs = new Dictionary<string, Word.Document>();
    const int batchLen = 20;
    private IProgressObservable progressListener;
    ExcelDataSource? excelDataSource;


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

    public void Start(JsonSettings settings)
        {
        //progressForm.StopRequested = true;

        //do
        //    {
        //    do
        //        {
        //        System.Windows.Forms.Application.DoEvents();

        //        if (!progressForm.StopRequested)
        //            break;

        //        } while (true);

        //    } while (!SendAllDocs());
        //    } while (!MergeModifySaveAll(CInt(progressForm.txtStart.value)));
        //this.TestExcel();
        //if (!this.PerformChecks())
        //    return;
        //this.SaveJsonSettings();

        //MergeModifySaveAllAsync(0);
        //SendAllDocs(settings);
        //DocBuilder docBuilder = new DocBuilder(settings.WordTemplateFileName);
        if (!this.PerformChecks(settings))
            return;
        if (settings.NamedRange == null)
            return;
        if (this.excelDataSource is null)
            return;
        if (this.excelDataSource is null)
            return;
        this.progressListener.ReportInfo("Preparing data...");
        AllowUIToUpdate();
        var data = this.excelDataSource.GetData(settings.NamedRange);
        this.excelDataSource.CloseExcel();
        var docBuilder = new DocBuilder(settings.WordTemplateFileName, data);
        var checkResults = docBuilder.GetChecksResults();
        if(checkResults.Count > 0)
            {
            this.progressListener.ReportError(String.Join("\n", checkResults));
            docBuilder.CloseAll();
            return;
            }
        this.progressListener.ReportInfo("Creating mail documents...");
        docBuilder.BuildDoc(1);
        docBuilder.BuildDoc(2);
        //docBuilder.BuildDoc(3);
        //docBuilder.BuildDoc(4);
        //docBuilder.BuildDoc(5);
        this.progressListener.ReportInfo("Sending mails...");
        Thread.Sleep(1000); //probably not needed.
        MailSender mailSender = new();
        mailSender.SetProgressObservable(this.progressListener);
        mailSender.SendAllDocs(settings);
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

    private int GetAccountIndex(string emailAddress)
        {
        for (int i = 1; i <= this.GetOutlook().Session.Accounts.Count; i++)
            {
            if (this.GetOutlook().Session.Accounts[i].DisplayName == emailAddress)
                return i;
            }

        return -1;
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
