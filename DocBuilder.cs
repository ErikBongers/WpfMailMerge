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

internal class DocBuilder
    {
    private readonly string sourceDocPath;
    private readonly Word.Application word;
    private readonly ExcelData excelData;
    private string mailMergeTempDir;
    private string mergedDocsDir;
    private string templateDocPath;

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
                }
            doc.SaveAs2(FileName: this.templateDocPath, AddToRecentFiles: false);
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
            foreach (Word.Range storyRange in doc.StoryRanges)
                {
                for (int i = 1; i < excelData.Headers.Count; i++)
                    {
                    while (true)
                        {
                        Word.Range searchRange = storyRange.Duplicate;
                        Word.Find find = searchRange.Find;
                        find.Text = $"{{{{{i}}}}}";
                        find.Execute();
                        if (!find.Found)
                            break;
                        searchRange.Text = excelData.GetRow(rowIndex)[i];
                        }
                    }
                while (true) 
                    {
                    Word.Range collapseRangeStart = storyRange.Duplicate;
                    Word.Find findStart = collapseRangeStart.Find;
                    findStart.Text = $"%%COLLAPSE%%";
                    findStart.Execute();
                    if (!findStart.Found)
                        break;
                    Word.Range collapseRangeEnd = storyRange.Duplicate;
                    Word.Find findEnd = collapseRangeEnd.Find;
                    findEnd.Text = $"%%END COLLAPSE%%";
                    findEnd.Execute();
                    if (!findEnd.Found)
                        break;//todo: error: collapse end marker not found
                    Word.Range inbetweenRange = doc.Range(collapseRangeStart.End, collapseRangeEnd.Start);
                    var text = inbetweenRange.Text;
                    text = text.Replace("\a", "");
                    if (string.IsNullOrWhiteSpace(text))
                        {
                        Word.Range rangeToDelete = doc.Range(collapseRangeStart.Start, collapseRangeEnd.End);
                        rangeToDelete.Delete();
                        }
                    else
                        {
                        collapseRangeStart.Delete();
                        collapseRangeEnd.Delete();
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

    ~DocBuilder()
        {
        this.word.Quit();
        }
    }