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
    private Dictionary<string, Word.Document> attachmentDocs = [];

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

    private void CheckIncludedDocs(Word.Document templateDoc)
        {
        List<int> attachememntIndices = [];
        foreach (Word.Range storyRange in templateDoc.StoryRanges)
            {
            var searchRange = storyRange.Duplicate;
            while (true)
                {
                var section = FindSection(searchRange, "%%INSERT ", "%%");
                if (section == null)
                    break;
                string? fieldMarker = section.InbetweenRange.Text;
                if (fieldMarker is null)
                    continue;
                string indexStr = fieldMarker.Replace("{{", "").Replace("}}", "").Trim();
                if (int.TryParse(indexStr, out int index))
                    {
                    attachememntIndices.Add(index);
                    }
                else
                    throw new Exception($"Invalid attachment field marker: {fieldMarker}");
                searchRange.Start = section.EndMarker.End;
                }
            }
        List<string> attachementFilePaths = [.. excelData.Rows.SelectMany(row =>
            {
                List<string> filePaths = [];
                foreach (int index in attachememntIndices)
                    {
                    if (index < row.Count)
                        {
                        string filePath = row[index];
                        if (!string.IsNullOrWhiteSpace(filePath))
                            filePaths.Add(filePath);
                        }
                    }
                return filePaths;
            }).Distinct()];
        foreach (string filePath in attachementFilePaths)
            {
            if (!File.Exists(filePath))
                throw new Exception($"Attachment file not found: {filePath}");
            attachmentDocs[filePath] = this.word.Documents.Open(filePath, Visible: false);
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
            //collect the ranges
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
            CheckIncludedDocs(doc);
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
                //for (int i = 1; i < excelData.Headers.Count; i++)
                //    {
                //    while (true)
                //        { 
                //        Word.Range searchRange = storyRange.Duplicate;
                //        Word.Find find = searchRange.Find;
                //        find.Text = $"{{{{{i}}}}}";
                //        find.Execute();
                //        if (!find.Found)
                //            break;
                //        searchRange.Text = excelData.GetRow(rowIndex)[i];
                //        }
                //    }
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
                        Word.Document? attachmentDoc = null;
                        if (!attachmentDocs.TryGetValue(fileName, out attachmentDoc))
                            continue;
                        section.OuterRange.FormattedText = attachmentDoc.Content.FormattedText;
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