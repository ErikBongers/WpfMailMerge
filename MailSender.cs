using System.IO;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace WpfMailMerge;

internal class MailSender
    {
    private Outlook.Application outlook;
    private Word.Application word;
    private int progressMaxValue;
    private IProgressObservable progressListener;
    private Outlook.Account outAccount;
    private Outlook.NameSpace session;
    private Outlook.Accounts accounts;
    private JsonSettings settings;

    public MailSender(JsonSettings settings)
        {
        this.settings = settings;
        this.word = new Word.Application();
        this.outlook = new Outlook.Application();
        this.session = this.outlook.Session;
        this.accounts = session.Accounts;
        this.outAccount = accounts[settings.MailAccountIndex];
        this.progressListener = new DummyProgressObservable();
        }

    public void SetProgressObservable(IProgressObservable progressListener) => this.progressListener = progressListener;

    public bool SendAllDocs(CancellationToken cancelToken, IProgress<DocServer.Status> progress)
        {
        var files = Directory.GetFiles(DocBuilder.MergedDocsDir);

        this.progressMaxValue = files.Length;
        this.progressListener.ReportProgress(0, this.progressMaxValue);

        int progressCnt = 1;

        foreach (var file in files)
            {
            SendOneMail(file);
            this.progressListener.SetProgress(progressCnt);
            progressCnt++;
            }
        this.progressListener.ReportInfo("Finished sending mails.");
        return true;
        }

    public void SendOneMail(string file)
        {
        Word.Document doc = this.word.Documents.Open(file, ReadOnly: true, Visible: false);

        string[] recipients = doc.Variables[Constants.VAR_RECIPIENTS].Value.Split(';');
        string subject = doc.Variables[Constants.VAR_SUBJECT].Value;

        string[] attachments = [];

        try
            {
            attachments = doc.Variables[Constants.VAR_ATTACHMENTS].Value.Split(';');
            }
        catch { }

        doc.Content.Copy();
        Thread.Sleep(1000); //sometimes the clipboard is not ready yet, so we wait a bit

        Outlook.MailItem mailItem = (Outlook.MailItem)this.outlook.CreateItem(Outlook.OlItemType.olMailItem);

        mailItem.SendUsingAccount = outAccount;
        if (!string.IsNullOrEmpty(settings.OnBehalfOfEmail))
            mailItem.SentOnBehalfOfName = settings.OnBehalfOfEmail;

        if (settings.UseTestRecipient)
            {
            recipients = [settings.TestRecipient];
            }
        foreach (var recipient in recipients)
            {
            if (!string.IsNullOrEmpty(recipient))
                mailItem.Recipients.Add(recipient);
            }

        foreach (var att in attachments)
            {
            if (!string.IsNullOrEmpty(att))
                {
                mailItem.Attachments.Add(att);
                //string fileName = Path.GetFileName(att);
                //string neWDir = @"C:\NoSharePoint\Attachments";
                //string newPath = Path.Combine(neWDir, fileName);
                //mailItem.Attachments.Add(newPath);
                }
            }

        mailItem.Subject = subject;

        mailItem.ReplyRecipients.Add("Academie Berchem <academie.berchem.muziek.woord@stedelijkonderwijs.be>");

        var inspector = mailItem.GetInspector;
        Word.Document? mailDoc = inspector.WordEditor as Word.Document;

        mailDoc?.Content.Paste();

        mailItem.Display();
        mailItem.Send();

        doc.Close(false);
        }

    public void CloseAll()
        {
        this.word.Quit();
        Marshal.FinalReleaseComObject(this.word);
        Marshal.FinalReleaseComObject(this.accounts);
        Marshal.FinalReleaseComObject(this.session);
        this.outlook.Quit();
        Marshal.FinalReleaseComObject(this.outlook);
        }
    
    ~MailSender()
        {
        CloseAll();
        }
    }


