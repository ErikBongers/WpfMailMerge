
namespace WpfMailMerge.Progress;

internal enum StatusType { Message, Progress, Error, ExcelRanges, ExcelData,
    Finished
    }

internal enum Cmd { CreateTempate, GenerateDocs }

internal class ProgressInfo
    {
    public int StartValue;
    public int MaxValue;
    public int CurrentValue;
    }

internal class Error
    {
    public required string message;
    public int code;
    }

internal class Status
    {
    public readonly StatusType StatusType;
    private readonly object Data;

    public Status(string message) { StatusType = StatusType.Message; Data = message; }
    public Status(string errorMessage, int errorCode) { StatusType = StatusType.Error; Data = new Error {message = errorMessage, code = errorCode }; }
    public Status(List<RangeDef> rangeDefs) { StatusType = StatusType.ExcelRanges; Data = rangeDefs; }
    public Status(ExcelData excelData) { StatusType = StatusType.ExcelData; Data = excelData; }
    public Status(bool finishedWithoutErrors) { StatusType = StatusType.Finished; Data = finishedWithoutErrors; }

    public Status(int startValue, int maxValue, int currentValue)
        {
        StatusType = StatusType.Progress;
        Data = new ProgressInfo
            {
            StartValue = startValue,
            MaxValue = maxValue,
            CurrentValue = currentValue
            };
        }
    public string GetMessage() { return (string)this.Data; }
    public Error GetError() { return (Error)this.Data; }
    public ProgressInfo GetProgressInfo() { return (ProgressInfo)this.Data; }
    public List<RangeDef> GetExcelNamedRanges() { return (List<RangeDef>)this.Data; }
    public ExcelData GetExcelData() { return (ExcelData)this.Data; }
    public bool GetFinishedStatus() { return (bool)this.Data; }
    }



