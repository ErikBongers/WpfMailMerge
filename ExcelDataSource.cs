using System.Diagnostics;
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

        var excel = GetExcel();
        try
            {
            var workBook = excel.Workbooks.Open(this.dataSourceFileName);
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
            return ranges;
            }
        finally
            {
            this.CloseExcel();
            }
        }

    public void TestExcel()
        {
        var excel = GetExcel();
        try
            {
            var workBook = excel.Workbooks.Open(this.dataSourceFileName);
            foreach (Excel.Worksheet workSheet in workBook.Worksheets)
                {
                Debug.WriteLine($"Sheet Name: {workSheet.Name}");
                Debug.WriteLine($"SheetRange: {workSheet.UsedRange.Address}");

                foreach (Excel.ListObject excelTable in workSheet.ListObjects)
                    {
                    Debug.WriteLine($"Table Name: {excelTable.Name}");
                    Debug.WriteLine($"Table Range: {excelTable.Range.Address}");

                    // Example: Access data, e.g., print header
                    // excelTable.HeaderRowRange
                    }
                }

            //read all cells
            // Retrieve values into a 2D array
            Excel.Worksheet workSheet1 = workBook.Worksheets[1];
            Excel.Range usedRange = workSheet1.UsedRange;
            object[,] data = (object[,])usedRange.Value2;
            Debug.WriteLine(data[1, 1]);
            }
        finally
            {
            this.CloseExcel();
            }
        }

    public ExcelData GetData(string rangeName)
        {
        RangeDef? rangeDef = this.GetRanges().FirstOrDefault(r => r.DisplayName == rangeName);
        if (rangeDef == null)
            {
            throw new ArgumentException($"Range {rangeName} not found in Excel file.");
            }
        var excel = GetExcel();
        try
            {
            var workBook = excel.Workbooks.Open(this.dataSourceFileName);
            Excel.Range range;
            if (rangeDef.RangeType == RangeType.Sheet)
                {
                Excel.Worksheet workSheet = workBook.Worksheets[rangeDef.Name];
                range = workSheet.UsedRange;
                }
            else
                {
                Excel.Worksheet workSheet = workBook.Worksheets[rangeDef.Name];
                range = workSheet.ListObjects[rangeDef.Name].Range;
                }
            object[,] data = (object[,])range.Value2;
            return new ExcelData(data);
            }
        finally
            {
            this.CloseExcel();
            }
        }

    private Excel.Application GetExcel()
        {
        if (this.excel == null)
            {
            this.excel = new Excel.Application();
            }
        return this.excel;
        }

    private void CloseExcel()
        {
        if (this.excel != null)
            {
            this.excel.Quit();
            this.excel = null;
            }
        }

    ~ExcelDataSource()
        {
        if (this.excel != null)
            {
            this.excel.Quit();
            }
        }
    }


public class ExcelData
    {
    private object[,] data;
    public object[,] Data { get => data; private set => data = value; }

    public ExcelData(object[,] data)
        {
        this.data = data;
        }
    }