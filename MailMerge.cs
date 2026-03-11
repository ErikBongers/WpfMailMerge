using System;
using System.IO;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows;

public class MailMerge
    {
    //private FrmProgress progressForm;
    private Outlook.Application outlook;
    private Word.Application word;

    private string VAR_RECIPIENTS;
    private string VAR_ATTACHMENTS;
    private Dictionary<string, Word.Document> cachedWordDocs = new Dictionary<string, Word.Document>();

    public MailMerge()
        {
        //progressForm = new FrmProgress();
        outlook = new Outlook.Application();
        word = new Word.Application();

        VAR_RECIPIENTS = "Dko3Recepient";
        VAR_ATTACHMENTS = "Dko3Attachments";

        AddSenders();
        }

    private void AddSenders()
        {
        for (int i = 1; i <= outlook.Session.Accounts.Count; i++)
            {
            var acc = outlook.Session.Accounts[i];
            AddSender(acc.DisplayName, i);
            }
        }

    public void Start()
        {
        //progressForm.Show();
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
        MergeModifySaveAllAsync(0);
        }

    public void AddSender(string emailAddress, int index)
        {
        //progressForm.AddSender(emailAddress, index);
        }

    private int GetAccountIndex(string emailAddress)
        {
        for (int i = 1; i <= outlook.Session.Accounts.Count; i++)
            {
            if (outlook.Session.Accounts[i].DisplayName == emailAddress)
                return i;
            }

        return -1;
        }

    private bool MergeModifySaveAllAsync(int startIndex)
        {
        int batchLen = 50;
        string dataSourceFileName = @"C:\Users\erikb\Desktop\TestDataMailMergeV2.xlsm";
        string wordFileName = @"C:\Users\erikb\Desktop\MailMerge.docm";

        Word.Document mainDoc = word.Documents.Open(wordFileName, ReadOnly: false, Visible: true); //todo: readonly? invisible?

        int recCount = GetRecordCount(dataSourceFileName);

        if (recCount == -1)
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

        //progressForm.SetInfo("Starting mail merge...");
        //progressForm.MaxProgressValue = recCount;

        int currentBatchIdx = startIndex;

        while (currentBatchIdx < recCount) // && !progressForm.StopRequested)
            {
            bool result = MergeModifySaveRange(currentBatchIdx, batchLen);
            if (!result)
                return false;

            currentBatchIdx += batchLen;
            }

        //if (!progressForm.StopRequested)
        //    progressForm.SetInfo("Finished!");

        return true;
        }

    private bool MergeModifySaveRange(int startIndex, int rangeLen)
        {
        Word.Document mergedDoc;
        string dataSourceFileName = @"C:\Users\erikb\Desktop\TestDataMailMergeV2.xlsm";

        OpenDataSource(word, dataSourceFileName, startIndex, rangeLen);

        int recCount = word.ActiveDocument.MailMerge.DataSource.RecordCount;

        string lastFieldName =
            word.ActiveDocument.MailMerge.DataSource.DataFields[
                word.ActiveDocument.MailMerge.DataSource.DataFields.Count].Name;

        lastFieldName = lastFieldName.Replace("vestiging", "");
        int numberOfVestigingen = GetNumbers(lastFieldName, 0);

        for (int i = 1; i <= recCount; i++)
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

            //progressForm.SetProgress(fileIdxStr);

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
            mergedDoc.SaveAs2( $@"C:\Users\erikb\Desktop\Merged\File{fileIdxStr}.docx");
            mergedDoc.Close(false);
            }

        return true;
        }

    private void ExpandMailDoc(Word.Document doc, string recipients, string subject, string[] wordDocs, string[] attachments)
        {
        string[] arrRecipients =
            "erikbongers@outlook.com".Split(';');

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

        string strRecipients = string.Join(";", arrRecipients);
        doc.Variables.Add(VAR_RECIPIENTS, strRecipients);

        string strAtt =
            string.Join(";", attachments.Where(a => !string.IsNullOrEmpty(a)));

        doc.Variables.Add(VAR_ATTACHMENTS, strAtt);
        }

    private Word.Document getCachedDoc(string wordDoc)
        {
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

    private bool SendAllDocs()
        {
        //int senderIndex = GetAccountIndex(progressForm.SenderAccount);
        int senderIndex = 0;

        Outlook.Account outAccount = outlook.Session.Accounts[senderIndex];

        string dirName = @"C:\Users\erikb\Desktop\Merged\";

        var files = Directory.GetFiles(dirName);

        //progressForm.MaxProgressValue = files.Length;

        int i = 1;

        foreach (var file in files)
            {
            //System.Windows.Forms.Application.DoEvents();

            //if (progressForm.StopRequested)
            //    {
            //    progressForm.SetInfo("Stopped by user");
            //    return false;
            //    }

            //progressForm.SetProgress(i);

            Word.Document doc = word.Documents.Open(file, ReadOnly: true, Visible: false);

            string[] recipients = doc.Variables[VAR_RECIPIENTS].Value.Split(';');

            string[] attachments = new string[] { };

            try
                {
                attachments = doc.Variables[VAR_ATTACHMENTS].Value.Split(';');
                }
            catch { }

            doc.Content.Copy();
            doc.Close(false);

            Outlook.MailItem mailItem = (Outlook.MailItem)outlook.CreateItem(Outlook.OlItemType.olMailItem);

            mailItem.SendUsingAccount = outAccount;

            foreach (var recipient in recipients)
                {
                if (!string.IsNullOrEmpty(recipient))
                    mailItem.Recipients.Add(recipient);
                }

            foreach (var att in attachments)
                {
                if (!string.IsNullOrEmpty(att))
                    mailItem.Attachments.Add(att);
                }

            mailItem.Subject = "subject";

            mailItem.ReplyRecipients.Add("Academie Berchem <academie.berchem.muziek.woord@stedelijkonderwijs.be>");

            var inspector = mailItem.GetInspector;
            Word.Document? mailDoc = inspector.WordEditor as Word.Document;

            mailDoc?.Content.Paste();

            mailItem.Display();
            mailItem.Send();

            i++;
            }

        return true;
        }
    }