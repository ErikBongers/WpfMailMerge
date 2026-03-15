using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.IO;
using System.Reflection.Metadata;
using Word = Microsoft.Office.Interop.Word;

namespace WpfMailMerge;

internal class FieldDef 
    {
    public required string Name { get; set; }
    public required int Index { get; set; }
    }

internal class FieldRangeDef
    {
    public required int Start { get; set; }
    public required int End { get; set; }
    public required int Index { get; set; }
    }

internal class DocBuilder
    {
    private readonly string sourceDocPath;
    private readonly Word.Application word;
    private readonly ExcelData excelData;
    private string mailMergeTempDir;
    private string mergedDocsDir;
    private string templateDocPath;
    private List<FieldRangeDef> fieldRanges = [];
    private Dictionary<string, Word.Document> insertDocs = [];

    public DocBuilder(string templateDocPath, ExcelData excelData)
        {
        this.sourceDocPath = templateDocPath;
        this.word = new Word.Application();
        this.excelData = excelData;
        this.mailMergeTempDir = Path.Combine(Path.GetTempPath(), Constants.APP_NAME);
        this.mergedDocsDir = MergedDocsDir;
        if (!Directory.Exists(this.mergedDocsDir))
            Directory.CreateDirectory(this.mergedDocsDir);
        this.templateDocPath = Path.Combine(mailMergeTempDir, "template.docx");
        CreateTemplateDoc();
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
                    {
                    indices.Add(index);
                    }
                else
                    throw new Exception($"Invalid {startMarker} field marker: {fieldMarker}");
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

    private List<string> GetUniqueColumnValues(List<int> indices)
        {
        return [.. excelData.Rows.SelectMany(row =>
            {
                List<string> values = [];
                foreach (int index in indices)
                    {
                    if (index < row.Count)
                        {
                        string value = row[index];
                        if (!string.IsNullOrWhiteSpace(value))
                            values.Add(value);
                        }
                    }
                return values;
            }).Distinct()];
        }

    private void CheckIncludedDocs(Word.Document templateDoc)
        {
        List<int> includedIndices = GetMarkerIndices(templateDoc, "%%INSERT ", "%%", remove: false);
        List<string> includedFilePaths = GetUniqueColumnValues(includedIndices);
        foreach (string filePath in includedFilePaths)
            {
            if (!File.Exists(filePath))
                throw new Exception($"File to include not found: {filePath}");
            insertDocs[filePath] = this.word.Documents.Open(filePath, Visible: false);
            }
        }

    private void CheckAndRemoveAttachmentMarkers(Word.Document templateDoc)
        {
        List<int> attachememntIndices = GetMarkerIndices(templateDoc, "%%ATTACH ", "%%", remove: true);
        List<string> attachementFilePaths = GetUniqueColumnValues(attachememntIndices);
        foreach (string filePath in attachementFilePaths)
            {
            if (!File.Exists(filePath))
                throw new Exception($"File to attach not found: {filePath}");
            }
        }

    public static string MergedDocsDir => Path.Combine(Path.GetTempPath(), Constants.APP_NAME, "Merged");

    public static bool IsMergeDirEmpty()
        {
        return !Directory.EnumerateFiles(MergedDocsDir).Any();
        }

    public static void ClearMergeDir()
        {
        foreach (string file in Directory.EnumerateFiles(MergedDocsDir))
            {
            File.Delete(file);
            }
        }

    public void BuildDoc(int rowIndex)
        {
        BuildOneDoc(this.word.Documents.Open(this.templateDocPath), rowIndex);
        }

private void CreateTemplateDoc()
        {
        Word.Document doc = this.word.Documents.Open(this.sourceDocPath);
        doc.SaveAs2(FileName: this.templateDocPath, AddToRecentFiles: false);
        doc.Close();
        doc = this.word.Documents.Open(this.templateDocPath);

        try
            {
            foreach (Word.Range storyRange in doc.StoryRanges)
                {
                for (int i = 1; i < excelData.Headers.Count; i++)
                    {
                    string fieldName = excelData.Headers[i];
                    Word.Find find = storyRange.Find;
                    find.Text = $"{{{{{fieldName}}}}}";
                    find.Replacement.Text = $"{{{{{i}}}}}";
                    find.Execute(Replace: Word.WdReplace.wdReplaceAll);
                    }
                //find INSERT markers
                }
            
            CheckIncludedDocs(doc);
            CheckAndRemoveAttachmentMarkers(doc);

            //collect the ranges AFTER all modifications. This ensures that the ranges are correct.
            foreach (Word.Range storyRange in doc.StoryRanges)
                {
                for (int i = 1; i < excelData.Headers.Count; i++)
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
                        var fieldRange = new FieldRangeDef { Start = searchRange.Start, End = searchRange.End, Index = i };
                        fieldRanges.Add(fieldRange);
                        searchRange.Collapse(WdCollapseDirection.wdCollapseEnd);
                        }
                    }
                }
            fieldRanges = fieldRanges.OrderByDescending(fr => fr.Start).ToList();
            doc.Save();
            }
        finally
            {
            doc.Close();
            }
        }

    private void BuildOneDoc(Word.Document templateDoc, int rowIndex)
        {
        string outputDocPath = Path.Combine(this.mergedDocsDir, $"{Constants.MERGED_FILE_PREFIX}{rowIndex}.docx");
        templateDoc.SaveAs2(FileName: outputDocPath, AddToRecentFiles: false); //todo: use FormattedText to copy content instead of saving template as new doc
        Word.Document doc = this.word.Documents.Open(FileName: outputDocPath, Visible: false);
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
                            continue;
                        section.OuterRange.FormattedText = insertDoc.Content.FormattedText;
                        }
                    }
                }
            doc.Save();
            }
        finally
            {
            doc.Close();
            }
        }

    private Section? FindSection(Word.Range searchRange, string startMarker, string endMarker)
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

    ~DocBuilder()
        {
        this.word.Quit();
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