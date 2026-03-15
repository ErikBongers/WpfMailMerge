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
    private readonly FieldDef[] fieldDefs;
    private readonly Word.Application word;

    public DocBuilder(string templateDocPath, FieldDef[] fieldDefs)
        {
        this.templateDocPath = templateDocPath;
        this.fieldDefs = fieldDefs;
        this.word = new Word.Application();

        }

    public void BuildDoc(string outputDocPath, object[] row)
        {

        }

    ~DocBuilder()
        {
        this.word.Quit();
        }
    }