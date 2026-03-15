using Word = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfMailMerge;

internal class FieldDef 
    {
    public required string Name { get; set; }
    public required int Index { get; set; }
    }

internal class DocBuilder
    {
    private readonly string templateDocPath;
    private readonly Word.Application word;
    private readonly ExcelData excelData;

    public DocBuilder(string templateDocPath, ExcelData excelData)
        {
        this.templateDocPath = templateDocPath;
        this.word = new Word.Application();
        this.excelData = excelData;
        }

    public void BuildDoc(string outputDocPath, int rowIndex)
        {
        Word.Document doc = this.word.Documents.Open(this.templateDocPath);
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
            doc.SaveAs2(@"C:\Users\erikb\Desktop\sdf.docx");
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