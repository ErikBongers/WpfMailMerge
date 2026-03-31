using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Markup;
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
    private IWordToEmailStrategy wordToEmail;

    public MailSender(JsonSettings settings, IWordToEmailStrategy wordToEmail)
        {
        this.settings = settings;
        this.wordToEmail = wordToEmail;
        this.word = new Word.Application();
        this.outlook = new Outlook.Application();
        this.session = this.outlook.Session;
        this.accounts = session.Accounts;
        this.outAccount = accounts[settings.MailAccountIndex];
        this.progressListener = new DummyProgressObservable();
        }

    public void SetProgressObservable(IProgressObservable progressListener) => this.progressListener = progressListener;

    public bool SendAllDocs(CancellationToken cancelToken, IProgress<Progress.Status> progress)
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
        //Word.Document doc = this.word.Documents.Open(file, ReadOnly: true, Visible: false);
        OpaqueDoc doc = this.wordToEmail.OpenDoc(this.word, file);

        string[] recipients = this.wordToEmail.GetRecipients(doc);
        string subject = this.wordToEmail.GetSubject(doc);

        string[] attachments = this.wordToEmail.GetAttachments(doc);

        this.wordToEmail.MaybeCopy(doc);

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
                }
            }

        mailItem.Subject = subject;

        mailItem.ReplyRecipients.Add("Academie Berchem <academie.berchem.muziek.woord@stedelijkonderwijs.be>");

        wordToEmail.FillEmail(doc, mailItem);
        mailItem.Send();

        this.wordToEmail.CloseDoc(doc);
        this.wordToEmail.MarkDocAsSent(doc, file);
        }

    public void CloseAll()
        {
        try { this.word.Quit(false); } catch (Exception) { } //word may already have been closed by the other thread.
        Marshal.FinalReleaseComObject(this.word);
        Marshal.FinalReleaseComObject(this.accounts);
        Marshal.FinalReleaseComObject(this.session);
        Marshal.FinalReleaseComObject(this.outlook);
        }
    
    ~MailSender()
        {
        CloseAll();
        }
    }


