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
    private List<PlaceHolderDef> placeHolders = [];
    private Dictionary<string, Word.Document> insertDocs = [];
    private List<FilesPlaceHolder> attachmentPlaceHolders = [];
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
            //do this BEFORE reversing!
            BuildSectionPlaceHolders(Constants.COLLAPSE_MARKER, Constants.END_COLLAPSE_MARKER);
            this.placeHolders = this.placeHolders.OrderByDescending(p => p.Pos).ToList();
            if (this.errors.Count > 0)
                return;
            CheckFiles();
            if (this.errors.Count > 0)
                return;
            var subjects = this.placeHolders
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
            var mailTos = this.placeHolders
                .Where(p => p is DecoratedStringPlaceHolder)
                .Cast<DecoratedStringPlaceHolder>()
                .Where(p => p.MarkerName == Constants.MAILTO_MARKER)
                .ToList();
            if (mailTos.Count == 0)
                this.errors.Add(ErrorDefs.MissingMarker(Constants.MAILTO_MARKER));

            this.mailToIndices = this.placeHolders
                .Where(p => p is DecoratedStringPlaceHolder)
                .Cast<DecoratedStringPlaceHolder>()
                .Where(p => p.MarkerName == Constants.MAILTO_MARKER)
                .ToList()
                .SelectMany(m => m.GetFieldIndices())
                .ToList();

            this.attachmentPlaceHolders = this.placeHolders
                .Where(p => p is FilesPlaceHolder)
                .Cast<FilesPlaceHolder>()
                .Where(p => p.MarkerName == Constants.ATTACH_MARKER)
                .ToList();

            if (this.errors.Count > 0)
                return;

            OpenIncludedFiles();

            this.templateDoc.SaveAs2(@"C:\Users\erikb\Desktop\test.docx");
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
                PlaceHolderDef? newPlaceHolder = null;
                string formattingKeepingText = "";
                if (section is null)
                    break;
                if (section.StartMarker.Text == "{{")
                    {
                    newPlaceHolder = new FieldPlaceHolder(section, this.excelData);
                    formattingKeepingText = "_";
                    }
                else //marker
                    {
                    var innerText = section.InbetweenRange.Text;
                    if (innerText.StartsWith(Constants.INSERT_MARKER) || innerText.StartsWith(Constants.ATTACH_MARKER))
                        newPlaceHolder = new FilesPlaceHolder(section, this.excelData.Headers, this.excelData);
                    else if (innerText.StartsWith(Constants.SUBJECT_MARKER))
                        newPlaceHolder = new DecoratedStringPlaceHolder(section, this.excelData.Headers, this.excelData);
                    else if (innerText.StartsWith(Constants.MAILTO_MARKER))
                        newPlaceHolder = new FieldsMarkerPlaceHolder(section, this.excelData.Headers, this.excelData);
                    else if (innerText.StartsWith(Constants.COLLAPSE_MARKER))
                        {
                        newPlaceHolder = new BeginSectionPlaceHolder(section);
                        formattingKeepingText = "_";
                        }
                    else if (innerText.StartsWith(Constants.END_COLLAPSE_MARKER))
                        {
                        newPlaceHolder = new EndSectionPlaceHolder(section);
                        formattingKeepingText = "_";
                        }
                    else
                        errors.Add(ErrorDefs.UnknownMarker(innerText.FirstWord()));
                    }
                if (newPlaceHolder?.Errors.Count > 0)
                    this.errors.AddRange(newPlaceHolder.Errors);
                if (newPlaceHolder != null)
                    this.placeHolders.Add(newPlaceHolder);
                searchRange = section.OuterRange;
                searchRange.Text = formattingKeepingText; // a Delete() or empty text will remove the formating of the original field or marker and may even trim spaces.
                searchRange.Collapse(Direction: WdCollapseDirection.wdCollapseStart);
                }
            }
        }
    
    private void BuildSectionPlaceHolders(string startMarker, string endMarker)
        {
        //outer scanner :: <other marker>* <section> <other marker>* <eof>
        //section :: <section start marke> <other marker>* <section end marker>
        //other marker :: // anything that is not a section begin or end marker
        //NOTE: every function starts at the current cursor already set.
        var cursor = new Cursor<PlaceHolderDef>(this.placeHolders.GetEnumerator());
        cursor.MoveNext();
        while (true)
            {
            SkipToSectionMarker(cursor, [startMarker, endMarker]);
            if (cursor.Current is null)
                break;//EOF
            ParseSection(cursor, startMarker, endMarker);
            }
        }

    private void ParseSection(Cursor<PlaceHolderDef> cursor, string startMarker, string endMarker)
        {
        if(cursor.Current is BeginSectionPlaceHolder beginHolder)
            {
            cursor.Eat();
            SkipToSectionMarker(cursor, [startMarker, endMarker]);
            ParseSection(cursor, startMarker, endMarker);
            if(cursor.Current is EndSectionPlaceHolder endHolder)
                {
                beginHolder.EndSectionPlaceHolder = endHolder;
                endHolder.BeginSectionPlaceHolder = beginHolder;
                cursor.MoveNext();
                }
            else
                {
                this.errors.Add(ErrorDefs.SectionWithoutEndMarker(startMarker));
                return;
                }
            }
        }

    private void SkipToSectionMarker(Cursor<PlaceHolderDef> cursor, List<string> sectionMarkers)
        {
        bool IsSectionMarker(PlaceHolderDef placeHolder)
            {
            if (placeHolder is MarkerPlaceHolder marker)
                return sectionMarkers.Contains(marker.MarkerName);
            return false;
            }

        cursor.Skip(IsSectionMarker);
        }

    public List<string> GetChecksResults()
        {
        return this.errors;
        }

    private void CheckFiles()
        {
        var filePlaceHolders = this.placeHolders
            .Where(p => p is FilesPlaceHolder)
            .Cast<FilesPlaceHolder>();
         

        List<string> includedFilePaths = this.excelData.GetUniqueColumnValues(filePlaceHolders, this.excelData);

        includedFilePaths
            .Where(path => Tools.FindAbsolutePath(path, this.excelData.GetDataDir()) is null)
            .ToList()
            .ForEach(path => this.errors.Add(ErrorDefs.FileToIncludeNotFound(path)));

        if (this.errors.Count > 0)
            return;

        foreach (string originalFilePath in includedFilePaths)
            {
            string generatedPath = Tools.FindAbsolutePath(originalFilePath, this.excelData.GetDataDir())!;
            this.absolutePaths[originalFilePath] = generatedPath;
            }
        }

    private void OpenIncludedFiles()
        {
        var filePlaceHolders = this.placeHolders
            .Where(p => p is FilesPlaceHolder)
            .Cast<FilesPlaceHolder>()
            .Where(p => p.MarkerName == Constants.INSERT_MARKER);

        List<string> includedFilePaths = this.excelData.GetUniqueColumnValues(filePlaceHolders, this.excelData);

        foreach (string originalFilePath in includedFilePaths)
            {
            insertDocs[originalFilePath] = this.documents.Open(this.absolutePaths[originalFilePath], Visible: false, ReadOnly: true);
            }
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

    public string BuildDoc(int rowIndex)
        {
        string fullName = Path.Combine(this.mergedDocsDir, $"{Constants.MERGED_FILE_PREFIX}{rowIndex}.docx");
        Word.Document doc = this.documents.Add(Visible: false);
        doc.Content.FormattedText = this.templateDoc.Content;
        try
            {
            int replacedUpToPos = doc.Content.End;
            foreach (var placeHolder in this.placeHolders)
                {
                if (placeHolder.Pos > replacedUpToPos)
                    continue; //skip placeholders until we are below replacedUpToPos 

                if (placeHolder is FieldPlaceHolder fieldPlaceHolder)
                    replacedUpToPos = fieldPlaceHolder.Replace(doc.Range(placeHolder.Pos, placeHolder.Pos + 1), excelData.GetRow(rowIndex), excelData); //todo: perhaps put the formattingPlaceholder (the "_") in the PlaceHolders.
                else if (placeHolder is FilesPlaceHolder filesPlaceHolder)
                    replacedUpToPos = filesPlaceHolder.Replace(doc.Range(placeHolder.Pos, placeHolder.Pos), excelData.GetRow(rowIndex), this.insertDocs);
                else if (placeHolder is BeginSectionPlaceHolder beginSection)
                    replacedUpToPos = beginSection.Replace(doc.Range(placeHolder.Pos, placeHolder.Pos + 1));
                else if (placeHolder is EndSectionPlaceHolder endSection)
                    replacedUpToPos = endSection.Replace(doc, this.placeHolders, excelData.GetRow(rowIndex));
                }

            //Add email variables (attachments, subject, mailto).
            List<string> attachments = [];
            foreach (var attachementPlaceHolder in this.attachmentPlaceHolders)
                {
                var values = attachementPlaceHolder.GetFieldValues(this.excelData.GetRow(rowIndex), this.excelData);
                foreach(var value in values)
                    if (!string.IsNullOrWhiteSpace(value))
                        attachments.Add(this.absolutePaths[value]);
                }
            attachments = attachments.Distinct().ToList();
            doc.Variables.Add(Constants.VAR_ATTACHMENTS, string.Join(";", attachments));
            if (this.subject is not null)
                {
                string subjectDecorated = this.subject.DecoratedString.Decorate(excelData.GetRow(rowIndex));
                doc.Variables.Add(Constants.VAR_SUBJECT, subjectDecorated);
                }

            var mailTos = string.Join(";", this.mailToIndices.Select(i => excelData.GetRow(rowIndex)[i]));

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
