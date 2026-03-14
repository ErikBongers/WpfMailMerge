using Microsoft.VisualBasic;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace WpfMailMerge;

public class MailAccount
    {
    public required string DisplayName { get; set; }
    public required int Index { get; set; }
    }

public class MailMerge : INotifyPropertyChanged
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
    private List<MailAccount> mailAccounts = new List<MailAccount>([new MailAccount { DisplayName="Loading...", Index=-1}]);
    public List<MailAccount> MailAccounts
        {
        get { return mailAccounts; }
        set { 
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
    private bool useTestRecipient = true;
    public bool UseTestRecipient
        {
        get { return useTestRecipient; }
        set { 
            useTestRecipient = value; OnPropertyChanged(nameof(UseTestRecipient)); 
            TestEmailVisibility = value ? Visibility.Visible : Visibility.Hidden;
            }
        }
    private string testRecipient = "erikbongers@outlook.com";
    public string TestRecipient
        {
        get { return testRecipient; }
        set { testRecipient = value; OnPropertyChanged(nameof(TestRecipient)); }
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
        set { wordTemplateFileName = value; OnPropertyChanged(nameof(WordTemplateFileName)); }
        }
    private string dataSourceFileName = @"C:\Users\erikb\Desktop\TestDataMailMergeV2.xlsm";
    public string DataSourceFileName
        {
        get { return dataSourceFileName; }
        set { dataSourceFileName = value; OnPropertyChanged(nameof(DataSourceFileName)); }
        }
    private string onBehalfOfEmail = "academie.berchem.muziek.woord@stedelijkonderwijs.be";
    public string OnBehalfOfEmail 
        {
        get { return onBehalfOfEmail; }
        set { onBehalfOfEmail = value; OnPropertyChanged(nameof(OnBehalfOfEmail)); }
        }
    public Visibility SenderComboVisibility {
        get { return this.MailAccounts.Count == 3 ? Visibility.Collapsed : Visibility.Visible; } //todo: test for 1 !!!
        }
    #endregion

    #region Constants
    private const string VAR_RECIPIENTS = "Dko3Recepient";
    private const string VAR_ATTACHMENTS = "Dko3Attachments";
    private const string VAR_SUBJECT = "Dko3Subject";
    private const string APP_NAME = "MailMerge";
    private const string SETTINGS_FILENAME = "settings.json";
    #endregion

    private Outlook.Application outlook;
    private Word.Application? word = null;
    private Word.Document? mainDoc;

    private Dictionary<string, Word.Document> cachedWordDocs = new Dictionary<string, Word.Document>();
    private int totalRecCount = -1;
    const int batchLen = 20;
    private readonly string mergedDocsDir;

    public MailMerge()
        {
        this.LoadJsonSettings();
        outlook = new Outlook.Application();
        this.MailAccounts = GetSenders();
        if (this.mailAccounts.Count == 1)
            {
            this.MailAccountIndex = this.mailAccounts[0].Index;
            }
        this.mergedDocsDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), APP_NAME, "Merged");
        if (!Directory.Exists(this.mergedDocsDir))
            Directory.CreateDirectory(this.mergedDocsDir);
        Debug.WriteLine($"Merged docs directory: {this.mergedDocsDir}");
        }

    private void LoadJsonSettings()
        {
        string localDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        string settingsFile = Path.Combine(localDir, APP_NAME, SETTINGS_FILENAME);
        if (File.Exists(settingsFile))
            {
            string json = File.ReadAllText(settingsFile);
            // Deserialize json to load settings
            JsonSettings? settings = System.Text.Json.JsonSerializer.Deserialize<JsonSettings>(json);
            if (settings == null)
                return;
            this.WordTemplateFileName = settings.WordTemplateFileName;
            this.DataSourceFileName = settings.DataSourceFileName;
            this.UseTestRecipient = settings.UseTestRecipient;
            this.TestRecipient = settings.TestRecipient;
            }
        }

    public void SaveJsonSettings()
        {
        string localDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        string appDir = Path.Combine(localDir, APP_NAME);
        if (!Directory.Exists(appDir))
            Directory.CreateDirectory(appDir);
        string settingsFile = Path.Combine(appDir, SETTINGS_FILENAME);
        JsonSettings settings = new JsonSettings
            {
            WordTemplateFileName = this.WordTemplateFileName,
            DataSourceFileName = this.DataSourceFileName,
            UseTestRecipient = this.UseTestRecipient,
            TestRecipient = this.TestRecipient
            };
        string json = System.Text.Json.JsonSerializer.Serialize(settings);
        File.WriteAllText(settingsFile, json);
        }


    private List<MailAccount> GetSenders()
        {
        List<MailAccount> accounts = new List<MailAccount>();
        for (int i = 1; i <= outlook.Session.Accounts.Count; i++)
            {
            accounts.Add(new MailAccount { DisplayName = outlook.Session.Accounts[i].DisplayName, Index = i });
            }
        return accounts;
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
            word?.Quit();
            }
        word = new Word.Application();
        return word;
        }

    public void Start()
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
        this.SaveJsonSettings();

        //MergeModifySaveAllAsync(0);
        SendAllDocs();
        }

    private void TestExcel()
    {
        var excel = new Excel.Application();
        try
        {
            var workBook = excel.Workbooks.Open(this.DataSourceFileName);
            foreach (Excel.Worksheet workSheet in workBook.Worksheets)
            {
                foreach (Excel.ListObject excelTable in workSheet.ListObjects)
                {
                    Console.WriteLine($"Table Name: {excelTable.Name}");
                    Console.WriteLine($"Table Range: {excelTable.Range.Address}");

                    // Example: Access data, e.g., print header
                    // excelTable.HeaderRowRange
                }
            }

            //read all cells
            // Retrieve values into a 2D array
            Excel.Worksheet workSheet1 = workBook.Worksheets[1];
            Excel.Range usedRange = workSheet1.UsedRange;
            object[,] data = (object[,])usedRange.Value2;
            Debug.WriteLine(data[1, 1]);
        }
        finally
        {
            excel.Quit();
        }
    }

    private bool PerformChecks()
        {
        if (!File.Exists(this.WordTemplateFileName))
            {
            MessageBox.Show($"Word template file not found: {this.WordTemplateFileName}");
            return false;
            }
        if (!File.Exists(this.DataSourceFileName))
            {
            MessageBox.Show($"Data source file not found: {this.DataSourceFileName}");
            return false;
            }
        //check if merged file folder is empty
        var files = Directory.GetFiles(this.mergedDocsDir, "*.docx", SearchOption.TopDirectoryOnly);
        if (files.Length > 0) {
            var reply = MessageBox.Show($"Merged documents directory is not empty: {this.mergedDocsDir}. Clear directory?", "Mail Merge", MessageBoxButton.YesNoCancel);
            if (reply == MessageBoxResult.Cancel)
                return false;
            if (reply == MessageBoxResult.Yes)
                {
                FileSystem.Kill(System.IO.Path.Combine(this.mergedDocsDir, "*.docx"));
                }
            }
        return true;
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
        resetWord();
        this.mainDoc = word?.Documents.Open(WordTemplateFileName, ReadOnly: false, Visible: true); //todo: readonly? invisible?
        this.totalRecCount = GetRecordCount(dataSourceFileName);

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

        this.StatusMessage = "Creating intermediate mail documents...";
        this.ProgressMaxValue = totalRecCount;
        this.setProgressInfo();

        int currentBatchIdx = startIndex;

        while (currentBatchIdx < totalRecCount) // && !progressForm.StopRequested)
            {
            bool result = MergeModifySaveRange(currentBatchIdx, batchLen);
            if (!result)
                return false;

            currentBatchIdx += batchLen;
            resetWord();
            mainDoc = word?.Documents.Open(WordTemplateFileName, ReadOnly: false, Visible: true); //todo: readonly? invisible?
            }

        //if (!progressForm.StopRequested)
        this.StatusMessage = "Finished!";

        return true;
        }

    private void setProgressInfo()
        {
        int percentage = (int)((double)ProgressValue / ProgressMaxValue * 100);
        this.ProgressInfo = $"{ProgressValue} ({percentage}%) of {ProgressMaxValue}";
        }

    private bool MergeModifySaveRange(int startIndex, int rangeLen)
        {
        if (word is null)
            return false;
        Word.Document mergedDoc;

        OpenDataSource(word, dataSourceFileName, startIndex, rangeLen);

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

            this.ProgressValue = int.Parse(fileIdxStr);
            this.setProgressInfo();

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
            mergedDoc.SaveAs2(Path.Combine(this.mergedDocsDir, $"File{fileIdxStr}.docx"));
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
        doc.Variables.Add(VAR_RECIPIENTS, recipients);
        string strAtt = string.Join(";", attachments.Where(a => !string.IsNullOrEmpty(a)));
        doc.Variables.Add(VAR_ATTACHMENTS, strAtt);
        doc.Variables.Add(VAR_SUBJECT, subject);
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

    private bool SendAllDocs()
        {
        if (word is null)
            this.word = this.resetWord();

        Outlook.Account outAccount = outlook.Session.Accounts[this.MailAccountIndex];

        var files = Directory.GetFiles(this.mergedDocsDir);

        this.StatusMessage = "Sending emails...";
        this.progressMaxValue = files.Length;

        int i = 1;

        foreach (var file in files)
            {
            //System.Windows.Forms.Application.DoEvents();

            //if (progressForm.StopRequested)
            //    {
            //    progressForm.SetInfo("Stopped by user");
            //    return false;
            //    }

            this.ProgressValue = i;

            Word.Document doc = word.Documents.Open(file, ReadOnly: true, Visible: false);

            string[] recipients = doc.Variables[VAR_RECIPIENTS].Value.Split(';');

            string[] attachments = new string[] { };

            try
                {
                attachments = doc.Variables[VAR_ATTACHMENTS].Value.Split(';');
                }
            catch { }

            doc.Content.Copy();
            Thread.Sleep(1000); //sometimes the clipboard is not ready yet, so we wait a bit

            Outlook.MailItem mailItem = (Outlook.MailItem)outlook.CreateItem(Outlook.OlItemType.olMailItem);

            mailItem.SendUsingAccount = outAccount;
            if (!string.IsNullOrEmpty(this.OnBehalfOfEmail))
                    mailItem.SentOnBehalfOfName = this.OnBehalfOfEmail;

            if (this.UseTestRecipient)
                {
                recipients = new string[] { this.TestRecipient };
                }
            foreach (var recipient in recipients)
                {
                if (!string.IsNullOrEmpty(recipient))
                    mailItem.Recipients.Add(recipient);
                }

            foreach (var att in attachments)
                {
                if (!string.IsNullOrEmpty(att)){
                    string fileName = Path.GetFileName(att);
                    string neWDir = @"C:\NoSharePoint\Attachments";
                    string newPath = Path.Combine(neWDir, fileName);
                    mailItem.Attachments.Add(newPath);
                }
                }

            mailItem.Subject = "subject";

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
