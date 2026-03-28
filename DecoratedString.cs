using System.Text.RegularExpressions;

namespace WpfMailMerge;

record Insert (int Pos, int Index);
public partial class DecoratedString
    {
    private readonly string orgText;
    private List<Insert> inserts = [];
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
                this.Errors.Add($"Can't find field {foundField}");
                continue;
                }
            this.templateText = this.templateText.Remove(pos, match.Length);
            this.inserts.Add(new Insert(pos, index));
            }
        this.inserts.Reverse();
        }


    public string Decorate(List<string> row) //todo: use enumerable instead of list here and everywhere possible.
        {
        string decorated = this.templateText;
        foreach (var insert in this.inserts)
            {
            string value = row[insert.Index];
            decorated = decorated.Insert(insert.Pos, value);
            }
        return decorated;
        }

    [GeneratedRegex(@"{{(.*?)}}")]
    private static partial Regex RxFieldPlaceHolder();
    }
