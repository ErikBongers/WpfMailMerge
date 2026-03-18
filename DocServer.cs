using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Channels;
using System.Threading.Tasks;
using System.Windows.Media.Animation;

namespace WpfMailMerge.DocServer;

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
    public Status(string errorMessage, int errorCode) { StatusType = StatusType.Message; Data = new Error {message = errorMessage, code = errorCode }; }

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

internal class DocServer
    {
    DocBuilder? docBuilder;
    private readonly string templateDocPath;
    private readonly ExcelData excelData;
    private readonly Channel<ProgressInfo> channel = Channel.CreateUnbounded<ProgressInfo>();

    public DocServer(string templateDocPath, ExcelData excelData)
        {
        this.templateDocPath = templateDocPath;
        this.excelData = excelData;
        }

    private void GenerateTemplate()
        {
        if (docBuilder != null)
            return;
        //template is created in constructor:
        docBuilder = new DocBuilder(this.templateDocPath, this.excelData);
        }

    }

