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

internal class FieldPlaceHolder : PlaceHolderDef
    {
    public IndexedFieldDef IndexedFieldDef;

    public FieldPlaceHolder(Section section, ExcelData excelData)
        : base(PlaceHolderType.Field, section)
        {
        (var indexedFieldDef , var error) = IndexedFieldDef.Create(section, excelData);
        this.IndexedFieldDef = indexedFieldDef;
        if (error is not null)
            this.Errors.Add(error);
        }
    
    public int Replace(Word.Range range, List<string> row, ExcelData excelData)
        {
        var value = this.IndexedFieldDef.GetValues(row, excelData);
        if(this.IndexedFieldDef.FieldDef.Options.Contains(Constants.MARKER_OPTION_EMPTYLINE))
            range.Text = string.Join("\n\n", value);
        else if(this.IndexedFieldDef.FieldDef.Options.Contains(Constants.MARKER_OPTION_NEWLINE))
            range.Text = string.Join("\n", value);
        else
            range.Text = string.Join(", ", value);
        return range.Start;
        }

    public override bool HasOnlyEmptyValues(List<string> row)
        {
        return row[this.IndexedFieldDef.Index].Trim().Length == 0;
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
    public string TextWithoutFields { get => this.DecoratedString.TemplateText; }
    public HashSet<string> Options { get; private set; }
    public string TextWithoutFieldsOrOptions { get; private set; }

    public DecoratedStringPlaceHolder(Section section, List<string> fieldNames, ExcelData excelData)
        : base(section)
        {
        this.DecoratedString = new DecoratedString(this.MarkerText, excelData);
        this.Errors.AddRange(this.DecoratedString.Errors);
        (this.Options, this.TextWithoutFieldsOrOptions) = Tools.ExtractOptions(this.TextWithoutFields, Constants.AllMarkerOptions);
        }

    public IEnumerable<int> GetFieldIndices()
        {
        return this.DecoratedString.Inserts.Select(i => i.IndexedFieldDef.Index);
        }
    
    public IEnumerable<string> GetFieldValues(List<string> row, ExcelData excelData)
        {
        return this.DecoratedString.Inserts.SelectMany(insert => {
            return insert.IndexedFieldDef.GetValues(row, excelData);
            });
        }

    public override bool HasOnlyEmptyValues(List<string> row)
        {
        return this.DecoratedString.Inserts.All(i => row[i.IndexedFieldDef.Index].Trim().Length == 0);
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

    public int Replace(Word.Range range, List<string> row, Dictionary<string, Word.Document> docs, ExcelData excelData)
        {
        var returnPos = range.Start;
        var insertList = this.DecoratedString.Inserts
            .OrderByDescending(i => i.Pos);
         
        foreach (var insert in insertList)
            {
            string[] fileNames = insert.GetValues(row, excelData);

            foreach (var fileName in fileNames)
                {
                Word.Document? insertDoc = null;
                if (!docs.TryGetValue(fileName, out insertDoc))
                    {
                    this.Errors.Add(ErrorDefs.CanNotOpenFile(fileName));
                    continue;
                    }
                range.FormattedText = insertDoc.Content.FormattedText;
                range.FormattedText = insertDoc.Content.FormattedText;
                range.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                }
            }
        return returnPos;
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