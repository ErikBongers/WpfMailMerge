using Microsoft.Office.Interop.Word;
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading.Tasks.Dataflow;
using System.Windows;
using Word = Microsoft.Office.Interop.Word;

namespace WpfMailMerge;

internal class FieldDef 
    {
    public required string Name { get; set; }
    public required int Index { get; set; }
    }

internal class PlaceholderDef
    {
    public required int Start { get; set; }
    public required int End { get; set; }
    public required int Index { get; set; }
    public required bool IsList { get; set; }
    }

internal enum PlaceHolderType { Field, Marker}
internal class NewPlaceHolderDef
    {
    public required PlaceHolderType Type { get; set; }
    public required string InnerText { get; set; }
    public required int Pos { get; set; }
    public required bool IsList { get; set; }
    public required string FieldName { get; set; }
    public required int ItemIndex { get; set; }
    }

internal record DocDef(string FullPath, Word.Document doc);

internal class DocBuilder
    {
    private readonly string sourceDocPath;
    private readonly Word.Application word;
    private readonly ExcelData excelData;
    private string mailMergeTempDir;
    private string mergedDocsDir;
    private string templateDocPath;
    private Word.Document templateDoc;
    private List<PlaceholderDef> fieldRanges = [];
    private List<NewPlaceHolderDef> newPlaceHolders = []; //todo: remove the 'new' once done.
    private List<MarkerDef> fileMarkerDefs = [];
    private Dictionary<string, Word.Document> insertDocs = [];
    private List<int> attachmentIndices = [];
    private DecoratedString? subject;
    private DecoratedString? mailTo;
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
            FirstScan();
            if (this.errors.Count > 0)
                return;
            CreatePlaceholderDefs();
            if (this.errors.Count > 0)
                return;
            var fieldNames = GetFieldNames();
            var missingFields = fieldNames.Except(excelData.Headers).ToList();
            if (missingFields.Count > 0)
                this.errors.Add($"Missing field(s): {string.Join(", ", missingFields)}");

            CheckMarkerFilesExist();
            CheckEmailMarkers(this.templateDoc);

            //collect the ranges AFTER all modifications. This ensures that the ranges are correct.
            foreach (Word.Range storyRange in this.templateDoc.StoryRanges)
                {
                for (int i = 0; i < excelData.Headers.Count; i++)
                    {
                    Word.Range searchRange = storyRange.Duplicate;
                    while (true)
                        {
                        Word.Find find = searchRange.Find;
                        find.ClearFormatting();
                        find.Forward = true;
                        find.Wrap = WdFindWrap.wdFindStop;
                        find.Text = $"{{{{{i}}}}}";
                        var found = find.Execute(Forward: true, Wrap: WdFindWrap.wdFindStop);
                        if (!found)
                            break;
                        var fieldRange = new PlaceholderDef { Start = searchRange.Start, End = searchRange.End, Index = i, IsList = false };
                        this.fieldRanges.Add(fieldRange);
                        searchRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                        }
                    }
                }
            this.fieldRanges = this.fieldRanges.OrderByDescending(fr => fr.Start).ToList();
            }
        finally
            {
            }
        }

    //collect placeholders that need to be removed from the template.
    private void FirstScan()
        {
        FirstScanForMarker("INSERT", false);
        FirstScanForMarker("ATTACH", true);
        }
    private void FirstScanForMarker(string marker, bool deleteMarker)
        {
        foreach (Word.Range storyRange in this.templateDoc.StoryRanges)
            {
            var searchRange = storyRange.Duplicate;
            while (true)
                {
                var placeHolder = FindNextPlaceHolder(searchRange);//todo: I'm only looking for %%, not {{
                if (placeHolder is null)
                    break;
                string innerText = placeHolder.InbetweenRange.Text;
                if (!innerText.Trim().StartsWith(marker + " "))
                    {
                    searchRange = placeHolder.OuterRange;
                    searchRange.Collapse(Direction: WdCollapseDirection.wdCollapseEnd);
                    continue;
                    }
                var fieldListString = innerText.Substring((marker + " ").Length);
                var fields = GetFieldsFromString(fieldListString);
                if (this.errors.Count > 0)
                    return;
                if (fields.Count == 0)
                    {
                    this.errors.Add($"No fields specified for %%{marker} marker.");
                    return;
                    }
                NewPlaceHolderDef placeHolderDef = new NewPlaceHolderDef
                    {
                    Type = PlaceHolderType.Marker,
                    Pos = placeHolder.OuterRange.Start,
                    ItemIndex = -1,
                    FieldName = "",
                    InnerText = placeHolder.OuterRange.Text, //todo: use InnerRange?
                    IsList = false
                    };
                MarkerDef markerDef = new MarkerDef { PlaceHolder = placeHolderDef, Fields = fields, MarkerTag = marker };
                this.fileMarkerDefs.Add(markerDef);

                searchRange = placeHolder.OuterRange;
                if (deleteMarker)
                    {
                    searchRange.Delete();
                    }
                else
                    {
                    searchRange.Text = $"%%{Constants.IDX_MARKER}{this.fileMarkerDefs.Count - 1}%%"; //todo: this replacement marker is not allowed in the original file.
                    searchRange.Collapse(Direction: WdCollapseDirection.wdCollapseEnd);
                    }
                }
            }
        }

    private void CreatePlaceholderDefs()
        {
        foreach (Word.Range storyRange in this.templateDoc.StoryRanges)
            {
            var searchRange = storyRange.Duplicate;
            while (true)
                {
                var placeHolder = FindNextPlaceHolder(searchRange);
                if (placeHolder is null)
                    break;
                var placeHolderText = placeHolder.OuterRange.Text;
                string innerText = placeHolder.InbetweenRange.Text;
                bool IsList = innerText.Contains('*'); //a bit loose - this doesn't check for position of the '*'.
                var placeHolderType = placeHolder.StartMarker.Text == "{{" ? PlaceHolderType.Field : PlaceHolderType.Marker;
                int itemIndex = -1;
                string fieldName = "";
                if (placeHolderType == PlaceHolderType.Field)
                    {
                    fieldName = innerText.Replace("*", "").Trim();
                    itemIndex = this.excelData.Headers.FindIndex(h => h == fieldName);
                    }
                else
                    {
                    if (innerText.StartsWith(Constants.IDX_MARKER))
                        {
                        itemIndex = int.Parse(innerText.Replace(Constants.IDX_MARKER, ""));
                        fieldName = Constants.IDX_MARKER;
                        }
                    else
                        fieldName = innerText; //todo: should probably be the first word only...but then there's "END COLLAPSE"...
                    }
                NewPlaceHolderDef placeholderDef = new NewPlaceHolderDef { Type = placeHolderType, InnerText = innerText, Pos = searchRange.Start, IsList = IsList, ItemIndex = itemIndex, FieldName = fieldName };
                this.newPlaceHolders.Add(placeholderDef);
                searchRange = placeHolder.OuterRange;
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

    private void CheckMarkerFilesExist()
        {
        List<string> includedFilesFieldNames = this.fileMarkerDefs.SelectMany(def => def.Fields).ToList();
        List<int> fieldIndexes = [];
        foreach (var fieldName in includedFilesFieldNames)
        {
            int index = this.excelData.Headers.IndexOf(fieldName);
            if(index < 0)
            {
                this.errors.Add($"Field {fieldName} not found."); //todo: perhaps move this check to the other fields check.
                continue;
            }
            fieldIndexes.Add(index);
        }
        List<string> includedFilePaths = this.excelData.GetUniqueColumnValues(fieldIndexes);

        foreach (string originalFilePath in includedFilePaths)
            {
            string? generatedPath = this.FindAbsolutePath(originalFilePath);
            if (generatedPath is null)
                {
                this.errors.Add($"File to include not found: {originalFilePath}");
                continue;
                }
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

    private void CheckEmailMarkers(Word.Document doc)
        {
        var subjects = GetMarkerValues(doc, "%%SUBJECT ", "%%", remove: true);
        if (subjects.Count > 1)
            this.errors.Add("Multiple SUBJECT markers found. Only one is allowed.");
        this.subject = new DecoratedString(subjects[0], this.excelData.Headers);
        if(this.subject.Errors.Count > 0)
            {
            this.errors.AddRange(this.subject.Errors);
            return;
            }
        
        var mailTos = GetMarkerValues(doc, "%%MAILTO ", "%%", remove: true);
        if (mailTos.Count > 1) //todo: allow multiple mailto markers for multiple recipients
            this.errors.Add("Multiple MAILTO markers found. Only one is allowed.");
        this.mailTo = new DecoratedString(mailTos[0], this.excelData.Headers);
        if (this.mailTo.Errors.Count > 0)
            {
            this.errors.AddRange(this.mailTo.Errors);
            return;
            }
        }


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
            foreach (var fieldRange in fieldRanges)
                {
                Word.Range range = doc.Range(fieldRange.Start, fieldRange.End);
                range.Text = excelData.GetRow(rowIndex)[fieldRange.Index];
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
                //find INSERT markers
                while (true)
                    {
                    var section = FindSection(storyRange, "%%INSERT ", "%%");
                    if (section == null)
                        break;
                    string? fileName = section.InbetweenRange.Text;
                    if (fileName is null || string.IsNullOrWhiteSpace(fileName))
                        {
                        section.OuterRange.Delete();
                        }
                    else
                        {
                        Word.Document? insertDoc = null;
                        if (!insertDocs.TryGetValue(fileName, out insertDoc))
                            continue; //todo: this is an error condition! At least report it.
                        section.OuterRange.FormattedText = insertDoc.Content.FormattedText;
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
                string subjectDecorated = this.subject.Decorate(excelData.GetRow(rowIndex));
                doc.Variables.Add(Constants.VAR_SUBJECT, subjectDecorated);
                }
            if (this.mailTo is not null)
                {
                string mailToDecorated = this.mailTo.Decorate(excelData.GetRow(rowIndex));
                doc.Variables.Add(Constants.VAR_RECIPIENTS, mailToDecorated);
                }
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


    private HashSet<string> GetFieldNames()
        {
        HashSet<string> fieldNames = [];
        foreach (var placeHolder in this.newPlaceHolders)
            {
            if (placeHolder.Type == PlaceHolderType.Field)
                fieldNames.Add(placeHolder.FieldName);
            }
        return fieldNames;
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

internal class MarkerDef
    {
    public required NewPlaceHolderDef PlaceHolder; //todo: probably not needed
    public List<string> Fields = [];
    public required string MarkerTag;
    }