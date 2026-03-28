using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace WpfMailMerge;

internal enum PlaceHolderType { Field, Marker }
internal abstract class PlaceHolderDef
    {
    public PlaceHolderType Type { get; private set; }
    public string InnerText { get; private set; }
    public int Pos { get; private set; }
    public List<string> Errors = [];

    public PlaceHolderDef(PlaceHolderType type, Section section)
        {
        this.Type = type;
        this.InnerText = section.InbetweenRange.Text;
        this.Pos = section.StartMarker.Start;
        }
    }

internal class FieldPlaceHolder : PlaceHolderDef
    {
    public bool IsList { get; private set; }
    public int FieldIndex { get; private set; }
    public string FieldName { get; private set; } //only for testing

    public FieldPlaceHolder(Section section, List<string> fieldNames)
        : base(PlaceHolderType.Field, section)
        {
        this.IsList = this.InnerText.Contains('*'); //a bit loose - this doesn't check for position of the '*'.
        string fieldName = this.InnerText.Replace("*", "").Trim();
        this.FieldName = fieldName; //test only
        this.FieldIndex = fieldNames.FindIndex(h => h == fieldName);
        if(this.FieldIndex == -1)
            this.Errors.Add(ErrorDefs.FieldNotFound(fieldName));
        }

    public int Replace(Word.Range range, List<string> row)
        {
        range.Text = row[this.FieldIndex];
        return range.Start;
        }
    }

internal abstract class MarkerPlaceHolder : PlaceHolderDef
    {
    public string MarkerName { get; private set; }
    public string MarkerText { get; private set; }

    public MarkerPlaceHolder(Section section)
        : base(PlaceHolderType.Marker, section)
        {
        var innerText = section.InbetweenRange.Text;
        this.MarkerName = innerText.Substring(0, innerText.IndexOf(" "));
        this.MarkerText = innerText.Substring(this.MarkerName.Length);
        }
    }

internal class DecoratedStringPlaceHolder : MarkerPlaceHolder
    {
    public DecoratedString DecoratedString { get; private set; }

    public DecoratedStringPlaceHolder(Section section, List<string> fieldNames)
        : base(section)
        {
        this.DecoratedString = new DecoratedString(this.MarkerText, fieldNames);
        this.Errors.AddRange(this.DecoratedString.Errors);
        }

    public IEnumerable<int> GetFieldIndices()
        {
        return this.DecoratedString.Inserts.Select(i => i.Index);
        }
    
    public IEnumerable<string> GetFieldValues(List<string> row)
        {
        return this.DecoratedString.Inserts.Select(i => row[i.Index]);
        }
    }

internal class FieldsMarkerPlaceHolder : DecoratedStringPlaceHolder
    {
    public FieldsMarkerPlaceHolder(Section section, List<string> fieldNames)
        : base(section, fieldNames)
        {
        if (this.DecoratedString.IsFieldsWithoutText())
            this.Errors.Add(ErrorDefs.OnlyFieldsWithoutTextExpected(section.OuterRange.Text));
        }
    }

internal class FilesPlaceHolder : FieldsMarkerPlaceHolder
    {
    public FilesPlaceHolder(Section section, List<string> fieldNames)
        : base(section, fieldNames) { }

    public int Replace(Word.Range range, List<string> row, Dictionary<string, Word.Document> docs)
        {
        var indexList = this.DecoratedString.Inserts
            .OrderByDescending(i => i.Pos)
            .Select(i => i.Index);
        foreach (var fieldIndex in indexList)
            {
            var fileName = row[fieldIndex];
            if (fileName == "")
                continue;

            Word.Document? insertDoc = null;
            if (!docs.TryGetValue(fileName, out insertDoc))
                { 
                this.Errors.Add(ErrorDefs.CanNotOpenFile(fileName));
                continue;
                }
            range.FormattedText = insertDoc.Content.FormattedText;
            range.Collapse(WdCollapseDirection.wdCollapseStart);
            }
        return range.Start;
        }
    }
