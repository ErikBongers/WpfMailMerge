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
    
    public List<RangeDef> GetRanges(string fileName)
        {
        if(this.ranges != null)
            {
            return this.ranges;
            }
        this.ranges = this.GetRangesForFile(fileName);
        return this.ranges;
        }
    
    private List<RangeDef> GetRangesForFile(string filePath)
        {
        if (this.excel == null)
            {
            throw new InvalidOperationException("Excel application is not initialized.");
            }
        var workBook = workbooks.Open(filePath);
        List<RangeDef> ranges = new List<RangeDef>();
        foreach (Excel.Worksheet workSheet in workBook.Worksheets)
            {
            ranges.Add(new RangeDef { BookName = filePath, Name = workSheet.Name, Range = workSheet.UsedRange.Address, RangeType = RangeType.Sheet });
            foreach (Excel.ListObject excelTable in workSheet.ListObjects)
                {
                ranges.Add(new RangeDef { BookName = filePath, Name = excelTable.Name, Range = excelTable.Range.Address, RangeType = RangeType.Table });
                }
            }
        workBook.Close(false);
        Marshal.FinalReleaseComObject(workBook);
        return ranges;
        }

    public ExcelData GetData(string filePath, string rangeName, bool mergeOtherExcels)
        {
        return GetDataInternal(filePath, rangeName, mergeOtherExcels, null)!;
        }

    private ExcelData? GetDataInternal(string filePath, string rangeName, bool mergeOtherExcels, ExcelData? matchHeadersFor)
        {
        var ranges = this.GetRanges(filePath);
        var rangeDef = ranges.FirstOrDefault(r => r.DisplayName.Equals(rangeName));
        if (rangeDef == null)
            {
            throw new ArgumentException($"Range not found in Excel file.");
            }
        var workBook = this.workbooks.Open(rangeDef.BookName);
        var workSheets = workBook.Worksheets;
        Excel.Worksheet workSheet;
        Excel.Range range;
        if (rangeDef.RangeType == RangeType.Sheet)
            {
            workSheet = workSheets[rangeDef.Name];
            range = workSheet.UsedRange;
            }
        else
            {
            workSheet = workBook.Worksheets[rangeDef.Name];
            range = workSheet.ListObjects[rangeDef.Name].Range;
            }
        object[,]? data;
        if (matchHeadersFor is not null)
            data = GetMatchingData(range, matchHeadersFor);
        else
            data = (object[,])range.Value2;
        Marshal.FinalReleaseComObject(range);
        Marshal.FinalReleaseComObject(workSheet);
        Marshal.FinalReleaseComObject(workSheets);
        workBook.Close(false);
        Marshal.FinalReleaseComObject(workBook);
        workBook = null;
        if (data is null)
            return null;
        var excelData = new ExcelData(data, rangeDef);
        //if (mergeOtherExcels)
        //    this.MergeOtherFiles(excelData);
        return excelData;
        }

    private object[,]? GetMatchingData(Excel.Range range, ExcelData matchHeadersFor)
        {
        //get first row and compare with headers.
        Excel.Range firstRow = range.Rows.Item[1];
        object[,] newHeaders = firstRow.Value2;
        throw new NotImplementedException("todo....");
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
            var ranges = this.GetRanges(file);
            foreach(var range in ranges)
                {
                var dataToMerge = this.GetDataInternal(range.BookName, range.Name, false, masterData);
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


public class ExcelData
    {
    private List<string> headers = new List<string>();
    private List<List<string>> rows = new List<List<string>>();
    public List<List<string>> Rows => rows;
    public readonly RangeDef rangeDef;

    public ExcelData(object[,] data, RangeDef rangeDef)
        {
        this.rangeDef = rangeDef;
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