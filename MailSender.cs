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

    public MailSender()
        {
        this.word = new Word.Application();
        this.outlook = new Outlook.Application();
        this.progressListener = new DummyProgressObservable();
        }

    public void SetProgressObservable(IProgressObservable progressListener) => this.progressListener = progressListener;

    public bool SendAllDocs(JsonSettings settings)
        {
        Outlook.NameSpace session;
        session = this.outlook.Session;
        var accounts = session.Accounts;
        Outlook.Account outAccount = accounts[settings.MailAccountIndex];

        var files = Directory.GetFiles(DocBuilder.MergedDocsDir);

        this.progressMaxValue = files.Length;
        this.progressListener.ReportProgress(0, this.progressMaxValue, "Sending emails...");

        int i = 1;

        foreach (var file in files)
            {
            //System.Windows.Forms.Application.DoEvents();

            //if (progressForm.StopRequested)
            //    {
            //    progressForm.SetInfo("Stopped by user");
            //    return false;
            //    }

            this.progressListener.SetProgress(i);

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

            i++;
            }
        this.progressListener.ReportInfo("Finished sending mails.");
        Marshal.FinalReleaseComObject(accounts);
        Marshal.FinalReleaseComObject(session);
        return true;
        }
    
    ~MailSender()
        {
        this.word.Quit();
        Marshal.FinalReleaseComObject(this.word);
        this.outlook.Quit();
        Marshal.FinalReleaseComObject(this.outlook);
        }
    }


