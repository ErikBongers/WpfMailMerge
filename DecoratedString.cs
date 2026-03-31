using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace WpfMailMerge;

public record class TextInsert (int Pos, IndexedFieldDef IndexedFieldDef)
    {
    public string GetValue(List<string> row, ExcelData excelData)
        {
        return this.IndexedFieldDef.GetValue(row, excelData);
        }
    }

public partial class DecoratedString
    {
    private readonly string orgText;
    public List<TextInsert> Inserts = [];
    public List<string> Errors = [];
    private string templateText;


    public DecoratedString(string text, ExcelData excelData)
        {
        this.orgText = text;
        this.templateText = text;
        this.ParseString(excelData);
        }

    private void ParseString(ExcelData excelData)
        {
        while(true)
            {
            Match match = RxFieldPlaceHolder().Match(this.templateText);
            if (!match.Success)
                break;
            string foundFieldDef = match.Groups[1].Value;
            FieldDef fieldDef = FieldDef.Parse(foundFieldDef);

            int pos = match.Index;
            this.templateText = this.templateText.Remove(pos, match.Length);
            (var indexedFieldDef, string? error) = IndexedFieldDef.Create(fieldDef, excelData);
            if (error is not null)
                this.Errors.Add(error);
            this.Inserts.Add(new TextInsert(pos, indexedFieldDef));
            }
        this.Inserts.Reverse();
        }

    public bool IsFieldsWithoutText()
        {
        return this.templateText.Trim().Length == 0;
        }

    public string Decorate(List<string> row)
        {
        string decorated = this.templateText;
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
