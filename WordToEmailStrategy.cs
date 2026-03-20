using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Runtime.ConstrainedExecution;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace WpfMailMerge;

internal interface IWordToEmailStrategy
    {
    void SaveDoc(Word.Document doc, string fullName);
    OpaqueDoc OpenDoc(Word.Application word, string fileName);
    void CloseDoc(OpaqueDoc doc);
    void MaybeCopy(OpaqueDoc doc);
    string[] GetRecipients(OpaqueDoc doc);
    string[] GetAttachments(OpaqueDoc doc);
    string GetSubject(OpaqueDoc doc);
    void FillEmail(OpaqueDoc doc, Outlook.MailItem mailItem);
    }

class OpaqueDoc
    {
    public required object doc;
    }

internal class WordCopyPaste : IWordToEmailStrategy
    {
    private JsonSettings settings;

    public WordCopyPaste(JsonSettings settings)
        {
        this.settings = settings;
        }

    public void SaveDoc(Document doc, string fullName)
        {
        doc.SaveAs2(FileName: fullName);
        }

    public OpaqueDoc OpenDoc(Word.Application word, string fileName)
        {
        return new OpaqueDoc { doc = word.Documents.Open(fileName, ReadOnly: true, Visible: false) };
        }
    
    public void CloseDoc(OpaqueDoc doc)
        {
        ((Word.Document)doc.doc).Close();
        }
    
    public string[] GetRecipients(OpaqueDoc doc)
        {
        return ((Word.Document)doc.doc).Variables[Constants.VAR_RECIPIENTS].Value.Split(';');
        }
    
    public string[] GetAttachments(OpaqueDoc doc)
        {
        string[] attachments = [];

        try
            {
            attachments = ((Word.Document)doc.doc).Variables[Constants.VAR_ATTACHMENTS].Value.Split(';');
            }
        catch { }
        return attachments;
        }

    public string GetSubject(OpaqueDoc doc)
        {
        return ((Word.Document)doc.doc).Variables[Constants.VAR_SUBJECT].Value;
        }

    public void MaybeCopy(OpaqueDoc doc)
        {
        ((Word.Document)doc.doc).Content.Copy();
        Thread.Sleep(settings.DelayAfterClipboardCopy); //sometimes the clipboard is not ready yet, so we wait a bit
        }

    public void FillEmail(OpaqueDoc doc, MailItem mailItem)
        {
        var inspector = mailItem.GetInspector;
        Word.Document? mailDoc = inspector.WordEditor as Word.Document;

        mailDoc?.Content.Paste();

        mailItem.Display();
        }
    }

internal class WordToRtfEmail : IWordToEmailStrategy
    {
    private JsonSettings settings;

    public WordToRtfEmail(JsonSettings settings)
        {
        this.settings = settings;
        }

    public void SaveDoc(Document doc, string fullName)
        {
        doc.SaveAs2(FileName: fullName + ".rtf", FileFormat: WdSaveFormat.wdFormatRTF);
        }

    public void MaybeCopy(OpaqueDoc doc) {}

    public void FillEmail(OpaqueDoc doc, MailItem mailItem)
        {
        mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatRichText;
        mailItem.RTFBody = System.Text.Encoding.ASCII.GetBytes((string)doc.doc);
        }

    public OpaqueDoc OpenDoc(Word.Application word, string fileName)
        {
        string rtf = File.ReadAllText(fileName+".rtf");
        return new OpaqueDoc { doc = rtf };
        }

    public void CloseDoc(OpaqueDoc doc)
        {
        //nothing to close - just a string.
        }
    

    public string[] GetRecipients(OpaqueDoc doc)
        {
        return getDocVar(((string)doc.doc), Constants.VAR_RECIPIENTS).Split(';');
        }

    public string[] GetAttachments(OpaqueDoc doc)
        {
        return getDocVar(((string)doc.doc), Constants.VAR_ATTACHMENTS).Split(';');
        }

    public string GetSubject(OpaqueDoc doc)
        {
        return getDocVar(((string)doc.doc), Constants.VAR_SUBJECT);
        }

    private string getDocVar(string rtf, string key)
        {
        //{\*\docvar {Dko3Recepient}{acercicek@gmail.com}}{\*\docvar {Dko3Subject}{Inschrijving voor Havin Masal ACER}}
        //{\*\docvar {Dko3Attachments}{C:\'5cUsers\'5cerikb\'5cDesktop\'5cDeMuziekschool.pdf}}
        // "C:\\'5cUsers\\'5cerikb\\'5cDesktop\\'5cDeMuziekschool.pdf"
        var index = rtf.IndexOf(key);
        if (index < 0)
            return "";
        var startOfVar = rtf.IndexOf('{', index + key.Length);
        if (startOfVar < 0)
            return "";
        var endOfVar = rtf.IndexOf('}', startOfVar);
        return rtf.Substring(startOfVar + 1, endOfVar - startOfVar - 1).Replace(@"\'5c", @"\");
        }
    }
