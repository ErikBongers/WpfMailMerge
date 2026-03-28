using System.Text.RegularExpressions;

namespace WpfMailMerge;

public record Insert (int Pos, int Index);
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
            string foundField = match.Groups[1].Value;
            int pos = match.Index;
            int index = fieldNames.IndexOf(foundField);
            if (index < 0)
                {
                this.Errors.Add($"Can't find field {foundField}.");
                continue;
                }
            this.templateText = this.templateText.Remove(pos, match.Length);
            this.Inserts.Add(new Insert(pos, index));
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
