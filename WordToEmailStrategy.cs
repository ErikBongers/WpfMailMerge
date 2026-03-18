using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Word;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace WpfMailMerge;

internal interface IWordToEmailStrategy
    {
    void SaveDoc(Word.Document doc);
    void MaybeCopy(Word.Document doc);
    void FillEmail(Word.Document doc, Outlook.MailItem mailItem);
    }

internal class WordCopyPaste : IWordToEmailStrategy
    {
    private JsonSettings settings;

    public WordCopyPaste(JsonSettings settings)
        {
        this.settings = settings;
        }

    public void SaveDoc(Document doc)
        {
        doc.Save();
        }

    public void MaybeCopy(Document doc)
        {
        doc.Content.Copy();
        Thread.Sleep(settings.DelayAfterClipboardCopy); //sometimes the clipboard is not ready yet, so we wait a bit
        }

    public void FillEmail(Document doc, MailItem mailItem)
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

    public void SaveDoc(Document doc)
        {
        doc.Save();
        doc.SaveAs2(FileName: doc.FullName + ".rtf", FileFormat: WdSaveFormat.wdFormatRTF);
        }

    public void MaybeCopy(Document doc) {}

    public void FillEmail(Document doc, MailItem mailItem)
        {
        mailItem.BodyFormat = Outlook.OlBodyFormat.olFormatRichText; 
        mailItem.RTFBody = System.Text.Encoding.ASCII.GetBytes(doc.FullName + ".rft");
        }
    }
