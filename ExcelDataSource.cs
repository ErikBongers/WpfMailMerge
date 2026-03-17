using Microsoft.Office.Interop.Word;
using System.Diagnostics;
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
    private readonly string dataSourceFileName;
    private Excel.Application? excel;
    private List<RangeDef>? ranges;

    public ExcelDataSource(string dataSourceFileName)
        {
        this.dataSourceFileName = dataSourceFileName;
        }

    
    public List<RangeDef> GetRanges()
        {
        if(this.ranges != null)
            {
            return this.ranges;
            }

        OpenExcel();
        if (this.excel == null)
            {
            throw new InvalidOperationException("Excel application is not initialized.");
            }
        var workbooks = excel.Workbooks;
        var workBook = workbooks.Open(this.dataSourceFileName);
        List<RangeDef> ranges = new List<RangeDef>();
        foreach (Excel.Worksheet workSheet in workBook.Worksheets)
            {
            ranges.Add(new RangeDef { Name = workSheet.Name, Range = workSheet.UsedRange.Address, RangeType = RangeType.Sheet });
            foreach (Excel.ListObject excelTable in workSheet.ListObjects)
                {
                ranges.Add(new RangeDef { Name = excelTable.Name, Range = excelTable.Range.Address, RangeType = RangeType.Table });
                }
            }
        this.ranges = ranges;
        workBook.Close(false);
        workbooks.Close();
        Marshal.FinalReleaseComObject(workBook);
        Marshal.FinalReleaseComObject(workbooks);
        return ranges;
        }

    public ExcelData GetData(string rangeName)
        {
        if (this.ranges is null)
            throw new ArgumentException("Ranges have not been set.");
        var rangeDef = this.ranges.FirstOrDefault(r => r.DisplayName.Equals(rangeName));
        if (rangeDef == null)
            {
            throw new ArgumentException($"Range not found in Excel file.");
            }
        OpenExcel();
        if (this.excel == null)
            {
            throw new InvalidOperationException("Excel application is not initialized.");
            }
        var workbooks = excel.Workbooks;
        var workBook = workbooks.Open(this.dataSourceFileName);
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
        object[,] data = (object[,])range.Value2;
        Marshal.FinalReleaseComObject(range);
        Marshal.FinalReleaseComObject(workSheet);
        Marshal.FinalReleaseComObject(workSheets);
        workBook.Close(false);
        workbooks.Close();
        Marshal.FinalReleaseComObject(workBook);
        Marshal.FinalReleaseComObject(workbooks);
        workBook = null;
        workbooks = null;
        return new ExcelData(data);
        }

    private void OpenExcel()
        {
        if (this.excel == null)
            {
            this.excel = new Excel.Application();
            }
        }

    public void CloseExcel()
        {
        if (this.excel is not null)
            {
            this.excel.Quit();
            Marshal.FinalReleaseComObject(this.excel);
            this.excel = null;
            GC.WaitForPendingFinalizers();
            GC.Collect();
            }
        }

    ~ExcelDataSource()
        {
        this.CloseExcel();
        }
    }


public class ExcelData
    {
    private List<string> headers = new List<string>();
    private List<List<string>> rows = new List<List<string>>();
    public List<List<string>> Rows => rows;

    public ExcelData(object[,] data)
        {
        for (int col = 1; col <= data.GetLength(1); col++)
            {
            headers.Add(data[1, col]?.ToString() ?? string.Empty);
            }
        for (int row = 2; row <= data.GetLength(0); row++)
            {
            List<string> rowData = new List<string>();
            for (int col = 1; col <= data.GetLength(1); col++)
                {
                rowData.Add(data[row, col]?.ToString() ?? string.Empty);
                }
            rows.Add(rowData);
            }
        }

    public List<string> Headers => headers;
    public List<string> GetRow(int index) => rows[index];
    }