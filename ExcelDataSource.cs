using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
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

    private ExcelData? GetMasterData(Excel.Workbook workBook, string rangeName, bool mergeOtherExcels)
        {
        var rangeDef = GetRangeDefFromName(workBook, rangeName);
        Excel.Range range = GetRangeFromDef(workBook, rangeDef);
        string? linkField = null;
        object[,]? data = (object[,])range.Value2;
        Marshal.FinalReleaseComObject(range);
        var excelData = new ExcelData(data, rangeDef, linkField);
        if (mergeOtherExcels)
            this.MergeOtherFiles(excelData);
        return excelData;
        }

    private ExcelData? GetLinkedData(Excel.Workbook workBook, string rangeName, ExcelData masterData)
        {
        var rangeDef = GetRangeDefFromName(workBook, rangeName);
        Excel.Range range = GetRangeFromDef(workBook, rangeDef);
        LinkedData? linkedData = GetMatchingData(range, masterData);
        Marshal.FinalReleaseComObject(range);
        if (linkedData is null)
            return null;
        var excelData = new ExcelData(linkedData.data, rangeDef, linkedData.linkField);
        return excelData;
        }

    private LinkedData? GetMatchingData(Excel.Range range, ExcelData masterData)
        {
        //get first row and compare with headers.
        Excel.Range firstRow = range.Rows.Item[1];
        object[,] newHeaders = firstRow.Value2;
        var newHeaderRow = ExcelData.ExtractHeaderRow(newHeaders);
        var unIndexedMasterHeaders = UnIndexedMasterHeaders(masterData.Headers);
        var matches = newHeaderRow.Intersect(unIndexedMasterHeaders).ToArray();
        if (matches is null)
            return null;
        int linkCount = matches.Count();
        if (linkCount > 1)
            throw new Exception("TODO: propery report too many link fields.");

        return new LinkedData((object[,])range.Value2, matches[0]);
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
        foreach (var file in files)
            {
            var workBook = this.workbooks.Open(file);
            var ranges = this.GetRangesForWorkbook(workBook);
            foreach(var range in ranges)
                {
                var dataToMerge = this.GetLinkedData(workBook, range.Name, masterData);
                if (dataToMerge is not null)
                    this.MergeFiles(masterData, dataToMerge);
                }
            workBook.Close(false);
            Marshal.FinalReleaseComObject(workBook);
            workBook = null;
            }
        }

    private void MergeFiles(ExcelData masterData, ExcelData otherData)
        {
        if (otherData.linkField is null)
            throw new Exception("todo: make this a compiler error: otherData should always have a linkField."); //class LinkedExcelData: ExcelData{...}
        //todo: Make hashtable inside LinkedExcelData
        foreach(var row in masterData.Rows)
            {
            //todo: get linked row, if any, and add the fields, headername separated by a dot.
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
    private List<string> headers = new List<string>();
    private List<List<string>> rows = new List<List<string>>();
    public List<List<string>> Rows => rows;
    public readonly RangeDef rangeDef;
    public string? linkField;

    public ExcelData(object[,] data, RangeDef rangeDef, string? linkField = null)
        {
        this.rangeDef = rangeDef;
        this.linkField = linkField;
        this.headers = ExtractHeaderRow(data);
        this.rows = ExtractBodyRows(data);
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
    }