
namespace WpfMailMerge.Progress;

internal enum StatusType { Message, Progress, Error }

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
    }



