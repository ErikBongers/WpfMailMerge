using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace WpfMailMerge;

internal class DocBuilder
    {
    private readonly string sourceDocPath;
    private readonly Word.Application word;
    private readonly ExcelData excelData;
    private string mailMergeTempDir;
    private string mergedDocsDir;
    private string templateDocPath;
    private Word.Document templateDoc;
    private List<PlaceHolderDef> newPlaceHolders = []; //todo: remove the 'new' once done.
    private Dictionary<string, Word.Document> insertDocs = [];
    private List<int> attachmentIndices = [];
    private List<int> mailToIndices = [];
    private DecoratedStringPlaceHolder? subject;
    private Word.Documents documents;
    private List<string> errors = [];
    private readonly IWordToEmailStrategy wordToEmail;

    public DocBuilder(string templateDocPath, ExcelData excelData, IWordToEmailStrategy wordToEmail)
        {
        this.sourceDocPath = templateDocPath;
        this.excelData = excelData;
        this.wordToEmail = wordToEmail;
        this.word = new Word.Application();
        this.documents = this.word.Documents;
        this.mailMergeTempDir = Path.Combine(Path.GetTempPath(), Constants.APP_NAME);
        this.mergedDocsDir = MergedDocsDir;
        if (!Directory.Exists(this.mergedDocsDir))
            Directory.CreateDirectory(this.mergedDocsDir);
        this.templateDocPath = Path.Combine(mailMergeTempDir, "template.docx");
        this.templateDoc = this.documents.Add(this.sourceDocPath) ?? throw new Exception("Template doc not found.");
        CreateTemplateDoc();
        }

    private void CreateTemplateDoc()
        {
        try
            {
            ExtractPlaceholders();
            this.newPlaceHolders = this.newPlaceHolders.OrderByDescending(p => p.Pos).ToList();
            if (this.errors.Count > 0)
                return;
            OpenIncludedFiles();
            if (this.errors.Count > 0)
                return;
            var subjects = this.newPlaceHolders
                .Where(p => p is DecoratedStringPlaceHolder)
                .Cast<DecoratedStringPlaceHolder>()
                .Where(p => p.MarkerName == Constants.SUBJECT_MARKER)
                .ToList();
            if (subjects.Count == 0)
                this.errors.Add(ErrorDefs.MissingMarker(Constants.SUBJECT_MARKER));
            if (subjects.Count > 1)
                this.errors.Add(ErrorDefs.MoreThanOneMarker(Constants.SUBJECT_MARKER));

            if (this.errors.Count > 0)
                return;

            this.subject = subjects[0];
            var mailTos = this.newPlaceHolders
                .Where(p => p is DecoratedStringPlaceHolder)
                .Cast<DecoratedStringPlaceHolder>()
                .Where(p => p.MarkerName == Constants.MAILTO_MARKER)
                .ToList();
            if (mailTos.Count == 0)
                this.errors.Add(ErrorDefs.MissingMarker(Constants.MAILTO_MARKER));

            this.mailToIndices = this.newPlaceHolders
                .Where(p => p is DecoratedStringPlaceHolder)
                .Cast<DecoratedStringPlaceHolder>()
                .Where(p => p.MarkerName == Constants.MAILTO_MARKER)
                .ToList()
                .SelectMany(m => m.GetFieldIndices())
                .ToList();

            this.attachmentIndices = this.newPlaceHolders
                .Where(p => p is FilesPlaceHolder)
                .Cast<FilesPlaceHolder>()
                .Where(p => p.MarkerName == Constants.ATTACH_MARKER)
                .ToList()
                .SelectMany(m => m.GetFieldIndices())
                .ToList();


            //this.templateDoc.SaveAs2(@"C:\Users\erikb\Desktop\test.docx");
            }
        finally
            {
            }
        }

    private void ExtractPlaceholders()
        {
        foreach (Word.Range storyRange in this.templateDoc.StoryRanges)
            {
            var searchRange = storyRange.Duplicate;
            while (true)
                {
                var section = FindNextPlaceHolder(searchRange);
                if (section is null)
                    break;
                if (section.StartMarker.Text == "{{")
                    this.newPlaceHolders.Add(new FieldPlaceHolder(section, this.excelData.Headers));
                else //marker
                    {
                    var innerText = section.InbetweenRange.Text;
                    if (innerText.StartsWith(Constants.INSERT_MARKER) || innerText.StartsWith(Constants.ATTACH_MARKER))
                        this.newPlaceHolders.Add(new FilesPlaceHolder(section, this.excelData.Headers));
                    else if (innerText.StartsWith(Constants.SUBJECT_MARKER))
                        this.newPlaceHolders.Add(new DecoratedStringPlaceHolder(section, this.excelData.Headers));
                    else if (innerText.StartsWith(Constants.MAILTO_MARKER))
                        this.newPlaceHolders.Add(new FieldsMarkerPlaceHolder(section, this.excelData.Headers));
                    else if (innerText.StartsWith(Constants.COLLAPSE_MARKER))
                        { } //todo this.newPlaceHolders.Add(new FieldsMarkerPlaceHolder(section, this.excelData.Headers));
                    else if (innerText.StartsWith(Constants.END_COLLAPSE_MARKER))
                        { } //todo this.newPlaceHolders.Add(new FieldsMarkerPlaceHolder(section, this.excelData.Headers));
                    else
                        errors.Add(ErrorDefs.UnknownMarker(innerText[..innerText.IndexOf(" ")]));
                    }
                searchRange = section.OuterRange;
                searchRange.Delete();
                searchRange.Collapse(Direction: WdCollapseDirection.wdCollapseStart);
                }
            }
        }
    
    private List<int> GetMarkerIndices(Word.Document templateDoc, string startMarker, string endMarker, bool remove)
        {
        List<int> indices = [];
        foreach (Word.Range storyRange in templateDoc.StoryRanges)
            {
            var searchRange = storyRange.Duplicate;
            while (true)
                {
                var section = FindSection(searchRange, startMarker, endMarker);
                if (section == null)
                    break;
                string? fieldMarker = section.InbetweenRange.Text;
                if (fieldMarker is null)
                    continue;
                string indexStr = fieldMarker.Replace("{{", "").Replace("}}", "").Trim();
                if (int.TryParse(indexStr, out int index))
                    indices.Add(index);
                else
                    {
                    this.errors.Add($"Invalid {startMarker} field marker: {fieldMarker}");
                    //fall through...
                    }
                if (remove)
                    {
                    section.OuterRange.Delete();
                    searchRange = storyRange.Duplicate;
                    }
                else
                    {
                    searchRange.Start = section.EndMarker.End;
                    }
                }
            }
        return indices;
        }

    public List<string> GetChecksResults()
        {
        return this.errors;
        }

    private List<string> GetMarkerValues(Word.Document templateDoc, string startMarker, string endMarker, bool remove) //todo: merge into GetMarkerIndices.
        {
        List<string> values = [];
        foreach (Word.Range storyRange in templateDoc.StoryRanges)
            {
            var searchRange = storyRange.Duplicate;
            while (true)
                {
                var section = FindSection(searchRange, startMarker, endMarker);
                if (section == null)
                    break;
                string? fieldMarker = section.InbetweenRange.Text;
                if (fieldMarker is null)
                    continue;
                values.Add(fieldMarker.Trim());
                if (remove)
                    {
                    section.OuterRange.Delete();
                    searchRange = storyRange.Duplicate;
                    }
                else
                    {
                    searchRange.Start = section.EndMarker.End;
                    }
                }
            }
        return values;
        }

    private void OpenIncludedFiles()
        {
        var fieldIndexes = this.newPlaceHolders
            .Where(p => p is FilesPlaceHolder)
            .Cast<FilesPlaceHolder>()
            .SelectMany(p => p.DecoratedString.Inserts.Select(i => i.Index));

        List<string> includedFilePaths = this.excelData.GetUniqueColumnValues(fieldIndexes);

        int errorCount = this.errors.Count;

        includedFilePaths
            .Where(path => this.FindAbsolutePath(path) is null)
            .ToList()
            .ForEach(path => this.errors.Add($"File to include not found: {path}"));

        if (errorCount > this.errors.Count)
            return;

        foreach (string originalFilePath in includedFilePaths)
            {
            string generatedPath = this.FindAbsolutePath(originalFilePath)!; //should already have been checked.
            this.absolutePaths[originalFilePath] = generatedPath;
            insertDocs[originalFilePath] = this.documents.Open(generatedPath, Visible: false, ReadOnly: true);
            }
        }

    private string? FindAbsolutePath(string originalFilePath)
        {
        string generatedPath = originalFilePath;
        if (!File.Exists(generatedPath))
            {
            var dataDir = this.excelData.GetDataDir();
            generatedPath = Path.Combine(dataDir, originalFilePath);
            if (!File.Exists(generatedPath))
                return null;
            }
        return generatedPath;
        }

    private Dictionary<string, string> absolutePaths = [];

    public static string MergedDocsDir => Path.Combine(Path.GetTempPath(), Constants.APP_NAME, "Merged");

    public static bool IsMergeDirEmpty()
        {
        return !Directory.EnumerateFiles(MergedDocsDir).Any();
        }

    public static bool ClearMergeDir()
        {
        bool hasErrors = false;
        foreach (string file in Directory.EnumerateFiles(MergedDocsDir))
        {
            try
            {
                File.Delete(file);
            }
            catch (Exception e)
            {
                hasErrors = true;
                Debug.WriteLine(e);
            }
        }
        if (hasErrors)
            MessageBox.Show("Can't delete all files!");
        return !hasErrors;
        }

    private List<string> GetFieldsFromString(string text)
    {
        int pos = 0;
        List<string> fields = new List<string>();
        while (true)
        {
            pos = text.IndexOf("{{", pos);
            if (pos == -1)
                break;
            int endPos = text.IndexOf("}}", pos);
            if (endPos == -1){
                this.errors.Add("Missing end of field delimiter '}}'.");
                return fields;
            }
            fields.Add(text.Substring(pos+2, endPos - pos-2));
            pos = endPos + 2;
        }
        return fields;
    }
    
    public string BuildDoc(int rowIndex)
        {
        string fullName = Path.Combine(this.mergedDocsDir, $"{Constants.MERGED_FILE_PREFIX}{rowIndex}.docx");
        Word.Document doc = this.documents.Add(Visible: false);
        doc.Content.FormattedText = this.templateDoc.Content;
        try
            {
            foreach (var placeHolder in this.newPlaceHolders)
                {
                Word.Range range = doc.Range(placeHolder.Pos, placeHolder.Pos);

                if (placeHolder is FieldPlaceHolder fieldPlaceHolder)
                    fieldPlaceHolder.Replace(range, excelData.GetRow(rowIndex));
                else if(placeHolder is FilesPlaceHolder filesPlaceHolder)
                    filesPlaceHolder.Replace(range, excelData.GetRow(rowIndex), this.insertDocs);
                }
            foreach (Word.Range storyRange in doc.StoryRanges)
                {
                while (true)
                    {
                    var section = FindSection(storyRange, "%%COLLAPSE%%", "%%END COLLAPSE%%");
                    if (section == null)
                        break;
                    var text = section.InbetweenRange.Text;
                    text = text.Replace("\a", "");
                    if (string.IsNullOrWhiteSpace(text))
                        {
                        section.OuterRange.Delete();
                        }
                    else
                        {
                        section.StartMarker.Delete();
                        section.EndMarker.Delete();
                        }
                    }
                }
            //Add email variables (attachments, subject, mailto).
            List<string> attachments = [];
            foreach (var idx in this.attachmentIndices)
                {
                string filePath = excelData.GetRow(rowIndex)[idx];
                if (!string.IsNullOrWhiteSpace(filePath))
                    attachments.Add(this.absolutePaths[filePath]);
                }
            attachments = attachments.Distinct().ToList();
            doc.Variables.Add(Constants.VAR_ATTACHMENTS, string.Join(";", attachments));
            if (this.subject is not null)
                {
                string subjectDecorated = this.subject.DecoratedString.Decorate(excelData.GetRow(rowIndex));
                doc.Variables.Add(Constants.VAR_SUBJECT, subjectDecorated);
                }

            var mailTos = string.Join(";", this.mailToIndices.Select(i => excelData.GetRow(i)));

            doc.Variables.Add(Constants.VAR_RECIPIENTS, string.Join(";", mailTos));
            this.wordToEmail.SaveDoc(doc, fullName);
            }
        finally
            {
            doc.Close();
            Marshal.FinalReleaseComObject(doc);
            }
        return fullName;
        }

    private static Section? FindSection(Word.Range searchRange, string startMarker, string endMarker)
        {
        Word.Range collapseRangeStart = searchRange.Duplicate;
        Word.Find findStart = collapseRangeStart.Find;
        findStart.Text = startMarker;
        findStart.Execute();
        if (!findStart.Found)
            return null;
        Word.Range collapseRangeEnd = collapseRangeStart.Duplicate;
        collapseRangeEnd.Collapse(WdCollapseDirection.wdCollapseEnd);
        Word.Find findEnd = collapseRangeEnd.Find;
        findEnd.Text = endMarker;
        findEnd.Execute();
        if (!findEnd.Found)
            return null; //todo: error: collapse end marker not found
        return new Section { StartMarker = collapseRangeStart.Duplicate, EndMarker = collapseRangeEnd.Duplicate };
        }
    
    private static Section? FindNextPlaceHolder(Word.Range searchRange)
        {
        Word.Range? fieldStart = null;
        Word.Range? markerStart = null;
        Word.Range rangeStart = searchRange.Duplicate;
        Word.Find findStart = rangeStart.Find;
        findStart.Text = "{{";
        findStart.Execute();
        if (findStart.Found)
            fieldStart = rangeStart.Duplicate;

        rangeStart = searchRange.Duplicate;
        findStart = rangeStart.Find;
        findStart.Text = "%%";
        findStart.Execute();
        if (findStart.Found)
            markerStart = rangeStart.Duplicate;

        if (fieldStart is null && markerStart is null)
            return null;

        bool useField = fieldStart is not null &&
                        (markerStart is null || fieldStart.Start < markerStart.Start);

        string endMarker = "%%";
        Word.Range sectionStartMarker;
        if (useField){
            sectionStartMarker = fieldStart!.Duplicate;
            endMarker = "}}";
            }
        else
            {
            sectionStartMarker = markerStart!.Duplicate;
            }

        //find endmarker (if required!)
        Word.Range collapseRangeEnd = sectionStartMarker.Duplicate;
        collapseRangeEnd.Collapse(WdCollapseDirection.wdCollapseEnd); //continue search where we ended.
        Word.Find findEnd = collapseRangeEnd.Find;
        findEnd.Text = endMarker;
        findEnd.Execute();
        if (!findEnd.Found)
            return null; //todo: error: end marker not found

        return new Section { StartMarker = sectionStartMarker.Duplicate, EndMarker = collapseRangeEnd.Duplicate };
        }

    public void CloseAll()
        {
        foreach (var doc in this.insertDocs.Values)
            {
            doc.Close();//todo: wrap in exception handler?
            Marshal.FinalReleaseComObject(doc);
            }
        this.insertDocs.Clear();
        try { this.templateDoc.Close(false); } catch (Exception e) { Debug.WriteLine("Can't close DocBuilder.templateDoc."); Debug.WriteLine(e); }
        Marshal.FinalReleaseComObject(this.templateDoc);
        Marshal.FinalReleaseComObject(this.documents);
        try { this.word.Quit(false); } catch (Exception) { } //word may already have been closed by the other thread.
        Marshal.FinalReleaseComObject(this.word);
        }

    ~DocBuilder()
        {
        this.CloseAll();
        }
    }

public class Section
    {
    public required Word.Range StartMarker { get; set; }
    public required Word.Range EndMarker { get; set; }
    public Word.Range InbetweenRange { 
        get
            {
            return this.StartMarker.Document.Range(this.StartMarker.End, this.EndMarker.Start);
            }
        }
    public Word.Range OuterRange
        {
        get
            {
            return this.StartMarker.Document.Range(this.StartMarker.Start, this.EndMarker.End);
            }
        }

    }
