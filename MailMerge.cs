using Microsoft.VisualBasic;
using System.Diagnostics;
using System.IO;
using System.Runtime;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace WpfMailMerge;

internal class MailMerge
    {
    private Outlook.Application? outlook;
    private Word.Application? word = null;
    private Word.Document? mainDoc;

    private Dictionary<string, Word.Document> cachedWordDocs = new Dictionary<string, Word.Document>();
    private int totalRecCount = -1;
    const int batchLen = 20;
    private IProgressObservable progressListener;
    private int progressMaxValue;
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

    private Word.Application resetWord()
        {
        if (this.word is not null)
            {
            foreach (var doc in cachedWordDocs.Values)
                {
                doc.Close(false);
                }
            cachedWordDocs.Clear();
            mainDoc?.Close(false);
            //word?.Quit();
            Marshal.ReleaseComObject(word);
            word = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            }
        word = new Word.Application();
        return word;
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
        var data = this.excelDataSource.GetData(settings.NamedRange);
        this.excelDataSource.CloseExcel();
        var docBuilder = new DocBuilder(settings.WordTemplateFileName, data);
        docBuilder.BuildDoc(1);
        docBuilder.BuildDoc(2);
        docBuilder.BuildDoc(3);
        docBuilder.BuildDoc(4);
        docBuilder.BuildDoc(5);
        //Thread.Sleep(1000); //probably not needed.
        //SendAllDocs(settings);
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

    private bool MergeModifySaveAllAsync(int startIndex, JsonSettings settings)
        {
        resetWord();
        this.mainDoc = word?.Documents.Open(settings.WordTemplateFileName, ReadOnly: false, Visible: true); //todo: readonly? invisible?
        this.totalRecCount = GetRecordCount(settings.DataSourceFileName);

        if (this.totalRecCount == -1)
            {
            string message = "No records found. Make sure the mail merge document is the active document or reset the mail merge source.";
            MessageBox.Show(message);
            return true;
            }

        //if (string.IsNullOrEmpty(progressForm.SenderAccount))
        //    {
        //    System.Windows.Forms.MessageBox.Show("Select a sender.");
        //    return false;
        //    }

        this.progressListener.ReportProgress(0, totalRecCount, "Creating intermediate mail documents...");

        int currentBatchIdx = startIndex;

        while (currentBatchIdx < totalRecCount) // && !progressForm.StopRequested)
            {
            bool result = MergeModifySaveRange(currentBatchIdx, batchLen, settings);
            if (!result)
                return false;

            currentBatchIdx += batchLen;
            resetWord();
            var documents = word?.Documents;
            mainDoc = documents?.Open(settings.WordTemplateFileName, ReadOnly: false, Visible: true); //todo: readonly? invisible?
            }

        return true;
        }

    private bool MergeModifySaveRange(int startIndex, int rangeLen, JsonSettings settings)
        {
        if (word is null)
            return false;
        Word.Document mergedDoc;

        OpenDataSource(word, settings.DataSourceFileName, startIndex, rangeLen);

        int batchRecCount = word.ActiveDocument.MailMerge.DataSource.RecordCount;

        string lastFieldName =
            word.ActiveDocument.MailMerge.DataSource.DataFields[
                word.ActiveDocument.MailMerge.DataSource.DataFields.Count].Name;

        lastFieldName = lastFieldName.Replace("vestiging", "");
        int numberOfVestigingen = GetNumbers(lastFieldName, 0);

        for (int i = 1; i <= batchRecCount; i++)
            {
            //System.Windows.Forms.Application.DoEvents();
            //if (progressForm.StopRequested)
            //    {
            //    progressForm.SetInfo("Stopped by user");
            //    return false;
            //    }

            var merge = word.ActiveDocument.MailMerge;

            merge.DataSource.FirstRecord = i;
            merge.DataSource.LastRecord = i;
            merge.DataSource.ActiveRecord = (Word.WdMailMergeActiveRecord)i;
            merge.Destination = Word.WdMailMergeDestination.wdSendToNewDocument;

            string email = merge.DataSource.DataFields["email"].Value;
            string naam = merge.DataSource.DataFields["naam"].Value;
            string voornaam = merge.DataSource.DataFields["voornaam"].Value;
            string fileIdxStr = merge.DataSource.DataFields["Idx"].Value;

            this.progressListener.SetProgress(int.Parse(fileIdxStr));

            string[] wordDocs = new string[numberOfVestigingen];
            string[] attachments = new string[numberOfVestigingen];

            for (int idx = 1; idx <= numberOfVestigingen; idx++)
                {
                wordDocs[idx - 1] = merge.DataSource.DataFields[$"vestiging{idx}_wordDoc"].Value;
                attachments[idx - 1] = merge.DataSource.DataFields[$"vestiging{idx}_bijlage"].Value;
                }

            merge.Execute(false);

            mergedDoc = word.ActiveDocument;
            ExpandMailDoc(mergedDoc, email, $"Start schooljaar voor {voornaam}", wordDocs, attachments);
            mergedDoc.SaveAs2(Path.Combine(DocBuilder.MergedDocsDir, $"File{fileIdxStr}.docx"));
            mergedDoc.Close(false);
            }

        return true;
        }

    private void ExpandMailDoc(Word.Document doc, string recipients, string subject, string[] wordDocs, string[] attachments)
        {
        foreach (string wordDoc in wordDocs)
            {
            if (!string.IsNullOrEmpty(wordDoc))
                {
                Word.Document insertDoc = getCachedDoc(wordDoc);

                Word.Range docRange = doc.Content;
                docRange.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                docRange.FormattedText = insertDoc.Content;
                }
            }
        doc.Variables.Add(Constants.VAR_RECIPIENTS, recipients);
        string strAtt = string.Join(";", attachments.Where(a => !string.IsNullOrEmpty(a)));
        doc.Variables.Add(Constants.VAR_ATTACHMENTS, strAtt);
        doc.Variables.Add(Constants.VAR_SUBJECT, subject);
        }

    private Word.Document getCachedDoc(string wordDoc)
        {
        if (word is null)
            throw new InvalidOperationException("Word application is not initialized.");
        Word.Document insertDoc;
        if (this.cachedWordDocs.ContainsKey(wordDoc))
            return this.cachedWordDocs[wordDoc];

        insertDoc = word.Documents.Open(wordDoc, ReadOnly: true, Visible: false);
        this.cachedWordDocs.Add(wordDoc, insertDoc);

        return insertDoc;
        }

    private int GetNumbers(string str, int occur)
        {
        var matches = Regex.Matches(str, @"(\d+)");
        return int.Parse(matches[occur].Value);
        }

    private void OpenDataSource(Word.Application word, string fileName, int startIdx, int idxCnt)
        {
        int maxIdx = startIdx + idxCnt;

        word.ActiveDocument.MailMerge.OpenDataSource(
            Name: fileName,
            ConfirmConversions: false,
            ReadOnly: false,
            LinkToSource: true,
            AddToRecentFiles: false,
            Format: Word.WdOpenFormat.wdOpenFormatAuto,
            Connection:
            $"Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source={fileName};Mode=Read;Extended Properties=\"HDR=YES;IMEX=1;\";",
            SQLStatement:
            $"SELECT * FROM `MAILMERGE$` where Idx >= {startIdx} and Idx < {maxIdx}",
            SubType: Word.WdMergeSubType.wdMergeSubTypeAccess
        );
        }

    private int GetRecordCount(string fileName)
        {
        if (word is null)
            throw new InvalidOperationException("Word application is not initialized.");
        word.ActiveDocument.MailMerge.OpenDataSource(
            Name: fileName,
            ConfirmConversions: false,
            ReadOnly: false,
            LinkToSource: true,
            AddToRecentFiles: false,
            Format: Word.WdOpenFormat.wdOpenFormatAuto,
            Connection:
            $"Provider=Microsoft.ACE.OLEDB.12.0;User ID=Admin;Data Source={fileName};Mode=Read;Extended Properties=\"HDR=YES;IMEX=1;\";",
            SQLStatement: "SELECT * FROM `MAILMERGE$`",
            SubType: Word.WdMergeSubType.wdMergeSubTypeAccess
        );

        return word.ActiveDocument.MailMerge.DataSource.RecordCount;
        }

    private bool SendAllDocs(JsonSettings settings)
        {
        if (word is null)
            this.word = this.resetWord();

        Outlook.Account outAccount = this.GetOutlook().Session.Accounts[settings.MailAccountIndex];

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

            Word.Document doc = word.Documents.Open(file, ReadOnly: true, Visible: false);

            string[] recipients = doc.Variables[Constants.VAR_RECIPIENTS].Value.Split(';');
            string subject = doc.Variables[Constants.VAR_SUBJECT].Value;

            string[] attachments = new string[] { };

            try
                {
                attachments = doc.Variables[Constants.VAR_ATTACHMENTS].Value.Split(';');
                }
            catch { }

            doc.Content.Copy();
            Thread.Sleep(1000); //sometimes the clipboard is not ready yet, so we wait a bit

            Outlook.MailItem mailItem = (Outlook.MailItem)this.GetOutlook().CreateItem(Outlook.OlItemType.olMailItem);

            mailItem.SendUsingAccount = outAccount;
            if (!string.IsNullOrEmpty(settings.OnBehalfOfEmail))
                mailItem.SentOnBehalfOfName = settings.OnBehalfOfEmail;

            if (settings.UseTestRecipient)
                {
                recipients = new string[] { settings.TestRecipient };
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

        return true;
        }
    }

internal class DummyProgressObservable : IProgressObservable
    {
    public void ReportProgress(int value, int maxValue, string info){}
    public void SetProgress(int value){}
    }
