using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace WpfMailMerge;

public enum PlaceHolderType { Field, Marker }
public abstract class PlaceHolderDef
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

    public virtual bool HasOnlyEmptyValues(List<string> row)  { return false; }
    }

public record FieldDef(string Name, bool IsList, string? SubFieldName);

internal class FieldPlaceHolder : PlaceHolderDef
    {
    public bool IsList { get; private set; }
    public int FieldIndex { get; private set; }
    public string FieldName { get; private set; } //only for testing
    public int? SubFieldIndex { get; private set; }
    public string? SubFieldName { get; private set; } //only for testing

    public FieldPlaceHolder(Section section, ExcelData excelData)
        : base(PlaceHolderType.Field, section)
        {
        FieldDef fieldDef = ParseFieldDef(this.InnerText);
        this.FieldName = fieldDef.Name;
        this.FieldIndex = excelData.Headers.FindIndex(h => h == this.FieldName);
        this.IsList = fieldDef.IsList; //todo: just include the FieldDef instead of copying all the values.
        if (this.FieldIndex == -1)
            this.Errors.Add(ErrorDefs.FieldNotFound(this.FieldName));
        if (fieldDef.SubFieldName is not null)
            {
            var linkedData = excelData.LinkedData[this.FieldIndex];
            var subIndex = linkedData.Headers.IndexOf(fieldDef.SubFieldName);
            if (this.FieldIndex == -1)
                {
                this.Errors.Add(ErrorDefs.FieldNotFound(fieldDef.SubFieldName)); //todo: make a more custom error for linked fields.
                return;
                }
            this.SubFieldName = fieldDef.SubFieldName;
            this.SubFieldIndex = subIndex;
            }
        }

    public int Replace(Word.Range range, List<string> row, ExcelData excelData)
        {
        if (this.SubFieldIndex is null)
            {
            range.Text = row[this.FieldIndex];
            return range.Start;
            }
        var key = row[this.FieldIndex];
        LinkedExcelData linkedData = excelData.LinkedData[this.FieldIndex];
        var linkedRow = linkedData.GetRow(key);
        var value = linkedRow[(int)this.SubFieldIndex];
        return range.Start;
        }

    public override bool HasOnlyEmptyValues(List<string> row)
        {
        return row[this.FieldIndex].Trim().Length == 0;
        }
    public static FieldDef ParseFieldDef(string text)
        {
        bool IsList = text.Contains('*'); //a bit loose - this doesn't check for position of the '*'.
        string fieldName = text.Replace("*", "").Trim();
        string? subFieldName = null;
        int sepPos = fieldName.IndexOf(Constants.SUBFIELD_SEPARATOR);
        if(sepPos >= 0)
            {
            subFieldName = fieldName.Substring(sepPos + 1);
            fieldName = fieldName.Substring(0, sepPos);
            }
        return new FieldDef(fieldName, IsList, subFieldName);
        }
    }

public abstract class MarkerPlaceHolder : PlaceHolderDef
    {
    public string MarkerName { get; private set; }
    public string MarkerText { get; private set; }

    public MarkerPlaceHolder(Section section)
        : base(PlaceHolderType.Marker, section)
        {
        var innerText = section.InbetweenRange.Text;
        this.MarkerName = innerText.FirstWord();
        this.MarkerText = innerText.Substring(this.MarkerName.Length);
        }
    }

public class DecoratedStringPlaceHolder : MarkerPlaceHolder
    {
    public DecoratedString DecoratedString { get; private set; }

    public DecoratedStringPlaceHolder(Section section, List<string> fieldNames, ExcelData excelData)
        : base(section)
        {
        this.DecoratedString = new DecoratedString(this.MarkerText, fieldNames, excelData);
        this.Errors.AddRange(this.DecoratedString.Errors);
        }

    public IEnumerable<int> GetFieldIndices()
        {
        return this.DecoratedString.Inserts.Select(i => i.Index);
        }
    
    public IEnumerable<string> GetFieldValues(List<string> row, ExcelData excelData)
        {
        return this.DecoratedString.Inserts.Select(insert => {
            if (insert.SubFieldIndex is null)
                return row[insert.Index];
            else
                {
                var key = row[insert.Index];
                var linkedData = excelData.LinkedData[insert.Index];
                var linkedRow = linkedData.GetRow(key);
                return linkedRow[(int)insert.SubFieldIndex];
                }
            });
        }

    public override bool HasOnlyEmptyValues(List<string> row)
        {
        return this.DecoratedString.Inserts.All(i => row[i.Index].Trim().Length == 0);
        }
    }

public class FieldsMarkerPlaceHolder : DecoratedStringPlaceHolder
    {
    public FieldsMarkerPlaceHolder(Section section, List<string> fieldNames, ExcelData excelData)
        : base(section, fieldNames, excelData)
        {
        if (!this.DecoratedString.IsFieldsWithoutText())
            this.Errors.Add(ErrorDefs.OnlyFieldsWithoutTextExpected(section.OuterRange.Text));
        }
    }

public class FilesPlaceHolder : FieldsMarkerPlaceHolder
    {
    public FilesPlaceHolder(Section section, List<string> fieldNames, ExcelData excelData)
        : base(section, fieldNames, excelData) { }

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

internal class BeginSectionPlaceHolder : MarkerPlaceHolder
    {
    public EndSectionPlaceHolder? EndSectionPlaceHolder;
    public BeginSectionPlaceHolder(Section section) : base(section)
        {
        }

    public IEnumerable<PlaceHolderDef> GetInnerPlaceHolders(IEnumerable<PlaceHolderDef> allPlaceHolders)
        {
        if (this.EndSectionPlaceHolder is null)
            return [];
        return allPlaceHolders.Where(p => p.Pos > this.Pos && p.Pos < this.EndSectionPlaceHolder.Pos);
        }
    
    public int Replace(Word.Range range)
        {
        range.Delete();
        return range.Start;
        }
    }

internal class EndSectionPlaceHolder : MarkerPlaceHolder
    {
    public BeginSectionPlaceHolder? BeginSectionPlaceHolder;
    public EndSectionPlaceHolder(Section section) : base(section)
        {
        }

    public IEnumerable<PlaceHolderDef> GetInnerPlaceHolders(IEnumerable<PlaceHolderDef> allPlaceHolders)
        {
        if (this.BeginSectionPlaceHolder is null)
            return [];
        return BeginSectionPlaceHolder.GetInnerPlaceHolders(allPlaceHolders);
        }

    public int Replace(Word.Document doc, IEnumerable<PlaceHolderDef> allPlaceHolders, List<string> row)
        {
        var innerPlaceHolders = this.GetInnerPlaceHolders(allPlaceHolders);
        bool hasValues = innerPlaceHolders.Any(p => !p.HasOnlyEmptyValues(row));
        if (hasValues)
            {
            Word.Range range = doc.Range(this.Pos, this.Pos+1); //+1 for placeholder!
            range.Delete();
            return this.Pos;
            }

        Word.Range collapseRange = doc.Range(this.BeginSectionPlaceHolder!.Pos+1, this.Pos+1); //first +1 is to NOT delete the Begin placeholder, the 2nd +1 is to delete the End placeholder.
        collapseRange.Delete();
        return this.BeginSectionPlaceHolder!.Pos+1; //+1 because the begin placeholder still has to be deleted.
        }
    }