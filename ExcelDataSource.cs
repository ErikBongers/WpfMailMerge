using System.Collections;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfMailMerge;

public enum RangeType
    {
    Sheet,
    Table,
    Waiting
    }

public class RangeDef
    {
    public required string BookName { get; set; }
    public required string SheetName { get; set; }
    public required string Name { get; set; }
    public required string Range { get; set; }
    public required RangeType RangeType { get; set; }
    public string DisplayName
        {
        get
            {
            return $"{this.RangeType}:{this.Name}";
            }
        }
    }

internal class ExcelDataSource
    {
    private IProgressObservable progressListener;

    private Excel.Application excel;
    private List<RangeDef>? ranges;
    private Excel.Workbooks workbooks;
    public List<string> warnings = [];
    public List<string> errors = [];

    public ExcelDataSource()
        {
        this.excel = new Excel.Application();
        this.workbooks = this.excel.Workbooks;
        this.progressListener = new DummyProgressObservable();
        }

    public void SetProgressObservable(IProgressObservable progressListener) => this.progressListener = progressListener;
    
    public List<RangeDef> GetRanges(string filePath)
        {
        if(this.ranges != null)
            {
            return this.ranges;
            }
        if (this.excel == null)
            {
            throw new InvalidOperationException("Excel application is not initialized.");
            }
        var workBook = workbooks.Open(filePath);
        this.ranges = this.GetRangesForWorkbook(workBook);
        workBook.Close(false);
        Marshal.FinalReleaseComObject(workBook);
        return this.ranges;
        }
    
    private List<RangeDef> GetRangesForWorkbook(Excel.Workbook workBook)
        {
        List<RangeDef> ranges = new List<RangeDef>();
        foreach (Excel.Worksheet workSheet in workBook.Worksheets)
            {
            ranges.Add(new RangeDef { BookName = workBook.FullName, SheetName = workSheet.Name, Name = workSheet.Name, Range = workSheet.UsedRange.Address, RangeType = RangeType.Sheet });
            foreach (Excel.ListObject excelTable in workSheet.ListObjects)
                {
                ranges.Add(new RangeDef { BookName = workBook.FullName, SheetName = workSheet.Name, Name = excelTable.Name, Range = excelTable.Range.Address, RangeType = RangeType.Table });
                }
            }
        return ranges;
        }
    
    private (List<RangeDef>, List<string>) GetFirstRangesForWorkbook(Excel.Workbook workBook)
        {
        List<string> warnings = [];
        List<RangeDef> ranges = new List<RangeDef>();
        foreach (Excel.Worksheet workSheet in workBook.Worksheets)
            {
            if (workSheet.ListObjects.Count > 1)
                warnings.Add($"Warning: {workBook}: sheet \"{workSheet.Name}\" has more than one table (or named range). Only the first is considered for linking.");
            if (workSheet.ListObjects.Count == 1){
                var excelTable = workSheet.ListObjects[1];
                ranges.Add(new RangeDef { BookName = workBook.FullName, SheetName = workSheet.Name, Name = excelTable.Name, Range = excelTable.Range.Address, RangeType = RangeType.Table });
                Marshal.FinalReleaseComObject(excelTable);
                }
            else
                ranges.Add(new RangeDef { BookName = workBook.FullName, SheetName = workSheet.Name, Name = workSheet.Name, Range = workSheet.UsedRange.Address, RangeType = RangeType.Sheet });
            }
        return (ranges, warnings);
        }

    public ExcelData GetData(string filePath, string rangeName, bool mergeOtherExcels)
        {
        return GetDataInternal(filePath, rangeName, mergeOtherExcels, null)!;
        }

    private ExcelData? GetDataInternal(string filePath, string rangeName, bool mergeOtherExcels, ExcelData? masterData)
        {
        var workBook = this.workbooks.Open(filePath);
        var excelData = GetMasterData(workBook, rangeName, mergeOtherExcels);
        workBook.Close(false);
        Marshal.FinalReleaseComObject(workBook);
        workBook = null;
        return excelData;
        }

    private RangeDef GetRangeDefFromName(Excel.Workbook workBook, string rangeName)
        {
        var ranges = this.GetRangesForWorkbook(workBook);
        var rangeDef = ranges.FirstOrDefault(r => r.DisplayName.Equals(rangeName));
        if (rangeDef == null)
            {
            throw new ArgumentException($"Range not found in Excel file.");
            }
        return rangeDef;
        }

    private Excel.Range GetRangeFromDef(Excel.Workbook workBook, RangeDef rangeDef)
        {
        Excel.Worksheet workSheet;
        Excel.Range range;
        if (rangeDef.RangeType == RangeType.Sheet)
            {
            workSheet = workBook.Worksheets[rangeDef.Name];
            range = workSheet.UsedRange;
            }
        else
            {
            workSheet = workBook.Worksheets[rangeDef.SheetName];
            range = workSheet.ListObjects[rangeDef.Name].Range;
            }

        Marshal.FinalReleaseComObject(workSheet);
        return range;
        }

    private ExcelData GetMasterData(Excel.Workbook workBook, string rangeName, bool mergeOtherExcels)
        {
        var rangeDef = GetRangeDefFromName(workBook, rangeName);
        Excel.Range range = GetRangeFromDef(workBook, rangeDef);
        object[,]? data = (object[,])range.Value2;
        Marshal.FinalReleaseComObject(range);
        var excelData = new ExcelData(data, rangeDef, this.warnings, this.errors);
        if (mergeOtherExcels)
            this.MergeOtherFiles(excelData);
        return excelData;
        }

    private LinkedExcelData? GetLinkedData(Excel.Workbook workBook, string rangeName, ExcelData masterData)
        {
        var rangeDef = GetRangeDefFromName(workBook, rangeName);
        Excel.Range range = GetRangeFromDef(workBook, rangeDef);
        LinkedData? linkedData = GetMatchingData(rangeDef, range, masterData);
        Marshal.FinalReleaseComObject(range);
        if (linkedData is null)
            return null;
        if (this.errors.Count > 0)
            return null;
        var excelData = new LinkedExcelData(linkedData.data, rangeDef, linkedData.linkField, this.warnings, this.errors);
        return excelData;
        }

    private LinkedData? GetMatchingData(RangeDef rangeDef, Excel.Range range, ExcelData masterData)
        {
        //get first row and compare with headers.
        Excel.Range firstRow = range.Rows.Item[1];
        object[,] newHeaders;
        if (firstRow.Columns.Count == 1){
            // create new object[1,1] but 1-based
            newHeaders = (Array.CreateInstance(typeof(object), new int[] { 1, 1 }, new int[] { 1, 1 }) as object[,])!;
            newHeaders[1,1] = firstRow.Value2;
            }
        else
            newHeaders = firstRow.Value2;
        var newHeaderRow = ExcelData.ExtractHeaderRow(newHeaders);
        var unIndexedMasterHeaders = UnIndexedMasterHeaders(masterData.Headers);
        var matches = newHeaderRow.Intersect(unIndexedMasterHeaders);
        int linkCount = matches.Count();
        if (linkCount == 0)
            return null;
        if (linkCount > 1){
            this.errors.Add($"Too many link fields in workbook {Path.GetFileName(rangeDef.BookName)} range {rangeDef.DisplayName}.");
            return null;
            }

        return new LinkedData((object[,])range.Value2, matches.First());
        }

    private static List<string> UnIndexedMasterHeaders(IEnumerable<string> headers)
        {
        return headers.Select(h => RemoveIndexes(h)).ToList();
        }

    private static string RemoveIndexes(string text)
        {
        int openBracketPos = text.IndexOf('[');
        if(openBracketPos >= 0)
            return text[..openBracketPos];
        return text;
        }

    private void MergeOtherFiles(ExcelData masterData)
        {
        string? dir = Path.GetDirectoryName(masterData.rangeDef.BookName);
        if (dir == null)
            throw new NotImplementedException("TODO: handle relative paths");
        var files = Directory.GetFiles(dir, "*.xls?");
        files = files.Where(file => file != masterData.rangeDef.BookName).ToArray();
        files = files.Where(file => !Path.GetFileName(file).StartsWith('~')).ToArray();
        foreach (var file in files)
            {
            var workBook = this.workbooks.Open(file);
            var (ranges, newWarnings) = this.GetFirstRangesForWorkbook(workBook);
            foreach(var range in ranges)
                {
                var dataToMerge = this.GetLinkedData(workBook, range.DisplayName, masterData);
                if (dataToMerge is not null)
                    {
                    this.MergeFiles(masterData, dataToMerge);
                    this.warnings = this.warnings.Concat(newWarnings).Where(w => w.Contains($"\"{range.Name}\"")).ToList(); //adding only warnings for the relavant sheet. todo: Assuming sheetname is between double quotes. Create a SheetWarning record.
                    }
                if (this.errors.Count > 0)
                    break;
                }
            workBook.Close(false);
            Marshal.FinalReleaseComObject(workBook);
            workBook = null;
            }
        }

    private void MergeFiles(ExcelData masterData, LinkedExcelData linkedData)
        {
        List<int> allMatchingKeyIndexes = [];
        foreach(var header in masterData.Headers.Select((text, i) => new {text, i}))
            {
            string baseHeader = RemoveIndexes(header.text);
            if (baseHeader == linkedData.linkField)
                allMatchingKeyIndexes.Add(header.i);
            }
        int linkedDataKeyIndex = linkedData.Headers.IndexOf(linkedData.linkField);
        //Add the extra header columns, except the key field.
        foreach (var masterIndex in allMatchingKeyIndexes)
            {
            var masterFieldName = masterData.Headers[masterIndex];
            foreach (var linkedFieldName in linkedData.Headers)
                if(linkedFieldName != linkedData.linkField)
                    masterData.Headers.Add(masterFieldName + "." + linkedFieldName);
            }
        foreach (var row in masterData.Rows)
            {
            foreach (var masterIndex in allMatchingKeyIndexes)
                {
                var linkedRow = linkedData.GetRow(row[masterIndex]);
                
                for (int i = 0; i < linkedRow.Count; i++)
                    if (i != linkedDataKeyIndex) //add fields except the key field.
                        row.Add(linkedRow[i]);

                }
            }
        }

    public void CloseAll()
        {
        this.workbooks.Close();
        Marshal.FinalReleaseComObject(this.workbooks);
        this.excel.Quit();
        Marshal.FinalReleaseComObject(this.excel);
        GC.WaitForPendingFinalizers();
        GC.Collect();
        }

    ~ExcelDataSource()
        {
        this.CloseAll();
        }
    }

public record LinkedData(object[,] data, string linkField);

public class ExcelData
    {
    protected List<string> headers = new List<string>();
    protected List<List<string>> rows = new List<List<string>>();
    public List<List<string>> Rows => rows;
    public readonly RangeDef rangeDef;
    public readonly List<string> warnings;
    public readonly List<string> errors;

    public ExcelData(object[,] data, RangeDef rangeDef, List<string> warnings, List<string> errors)
        {
        this.rangeDef = rangeDef;
        this.warnings = warnings;
        this.errors = errors;
        this.headers = ExtractHeaderRow(data);
        this.rows = ExtractBodyRows(data);
        }

    public string GetDataDir()
    {
        string fullPath = Path.GetFullPath(this.rangeDef.BookName);
        string? dir = Path.GetDirectoryName(fullPath);
        if (dir == null)
        {
            throw new Exception($"Can't find base directory for Workbook {this.rangeDef.BookName}");
        }
        return dir;
    }

    public static List<string> ExtractHeaderRow(object[,] data)
        {
        List<string> headers = new List<string>();
        for (int col = 1; col <= data.GetLength(1); col++)
            {
            headers.Add(data[1, col]?.ToString() ?? string.Empty);
            }
        return headers;
        }

    public static List<List<string>> ExtractBodyRows(object[,] data)
        {
        List<List<string>> rows = new List<List<string>>();
        for (int row = 2; row <= data.GetLength(0); row++)
            {
            List<string> rowData = new List<string>();
            for (int col = 1; col <= data.GetLength(1); col++)
                {
                rowData.Add(data[row, col]?.ToString() ?? string.Empty);
                }
            rows.Add(rowData);
            }
        return rows;
        }

    public List<string> Headers => headers;
    public List<string> GetRow(int index) => rows[index];
    public void Truncate(int max)
        {
        this.rows = this.rows.Take(max).ToList();
        }

    public List<string> GetUniqueColumnValues(IEnumerable<int> indices)
        {
        return [..this.Rows.SelectMany(row =>
            {
                List<string> values = [];
                foreach (int index in indices)
                    {
                    if (index < row.Count)
                        {
                        string value = row[index];
                        if (!string.IsNullOrWhiteSpace(value))
                            values.Add(value);
                        }
                    }
                return values;
            }).Distinct()];
        }
    }

class LinkedExcelData : ExcelData
    {
    public string linkField;
    private Dictionary<string, List<string>> dict = [];
    private List<string> nullRow;
    public LinkedExcelData(object[,] data, RangeDef rangeDef, string linkField, List<string> warnings, List<string> errors) 
        : base(data, rangeDef, warnings, errors)
        {
        this.linkField = linkField;
        this.FillDictionary();
        this.nullRow = this.headers.Select(h => "").ToList();
        }
    
    private void FillDictionary()
        {
        int keyIndex = this.headers.IndexOf(this.linkField);
        foreach(var row in this.Rows)
            {
            dict.Add(row[keyIndex], row); //todo: may fail with duplicate key -> report error.
            }
        }

    public List<string> GetRow(string key)
        {
        if (this.dict.ContainsKey(key))
            return this.dict[key];
        return this.nullRow;
        }
}
