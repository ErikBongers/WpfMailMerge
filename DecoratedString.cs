using System.Text.RegularExpressions;

namespace WpfMailMerge;

public record Insert (int Pos, int Index, bool IsList);
public partial class DecoratedString
    {
    private readonly string orgText;
    public List<Insert> Inserts = [];
    public List<string> Errors = [];
    private string templateText;


    public DecoratedString(string text, List<string> fieldNames)
        {
        this.orgText = text;
        this.templateText = text;
        this.ParseString(fieldNames);
        }

    private void ParseString(List<string> fieldNames)
        {
        while(true)
            {
            Match match = RxFieldPlaceHolder().Match(this.templateText);
            if (!match.Success)
                break;
            string foundFieldDef = match.Groups[1].Value;
            FieldDef fieldDef = FieldPlaceHolder.ParseFieldDef(foundFieldDef);

            int pos = match.Index;
            int index = fieldNames.IndexOf(fieldDef.Name);
            this.templateText = this.templateText.Remove(pos, match.Length);
            if (index < 0)
                {
                this.Errors.Add($"Can't find field {fieldDef.Name}.");
                continue;
                }
            this.Inserts.Add(new Insert(pos, index, fieldDef.IsList));
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
            string value = row[insert.Index];
            decorated = decorated.Insert(insert.Pos, value);
            }
        return decorated;
        }

    [GeneratedRegex(@"{{(.*?)}}")]
    private static partial Regex RxFieldPlaceHolder();
    }
