using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace WpfMailMerge;

internal class ExcelDataSource
    {
    private readonly string dataSourceFileName;
    private Excel.Application? excel;

    public ExcelDataSource(string dataSourceFileName)
        {
        this.dataSourceFileName = dataSourceFileName;
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


