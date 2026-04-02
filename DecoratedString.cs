using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace WpfMailMerge;

public record class TextInsert (int Pos, IndexedFieldDef IndexedFieldDef)
    {
    public string[] GetValues(List<string> row, ExcelData excelData)
        {
        return this.IndexedFieldDef.GetValues(row, excelData);
        }
    }

public partial class DecoratedString
    {
    private readonly string orgText;
    public List<TextInsert> Inserts = [];
    public List<string> Errors = [];
    public string TemplateText { get; private set ; }

    public DecoratedString(string text, ExcelData excelData)
        {
        this.orgText = text;
        this.TemplateText = this.ParseString(excelData, text);
        }

    private string ParseString(ExcelData excelData, string text)
        {
        while(true)
            {
            Match match = RxFieldPlaceHolder().Match(text);
            if (!match.Success)
                break;
            string foundFieldDef = match.Groups[1].Value;
            FieldDef fieldDef = FieldDef.Parse(foundFieldDef);

            int pos = match.Index;
            text = text.Remove(pos, match.Length);
            (var indexedFieldDef, string? error) = IndexedFieldDef.Create(fieldDef, excelData);
            if (error is not null)
                this.Errors.Add(error);
            this.Inserts.Add(new TextInsert(pos, indexedFieldDef));
            }
        this.Inserts.Reverse();
        return text;
        }

    public bool IsFieldsWithoutText()
        {
        return this.TemplateText.Trim().Length == 0;
        }

    public string Decorate(List<string> row)
        {
        string decorated = this.TemplateText;
        foreach (var insert in this.Inserts)
            {
            string value = row[insert.IndexedFieldDef.Index];
            decorated = decorated.Insert(insert.Pos, value);
            }
        return decorated;
        }

    [GeneratedRegex(@"{{(.*?)}}")]
    private static partial Regex RxFieldPlaceHolder();
    }
