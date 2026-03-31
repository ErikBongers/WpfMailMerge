using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfMailMerge;

public record class FieldDef(string Name, bool IsList, string? SubFieldName)
    {
    public static FieldDef Parse(string text)
        {
        bool IsList = text.Contains('*'); //a bit loose - this doesn't check for position of the '*'.
        string fieldName = text.Replace("*", "").Trim();
        string? subFieldName = null;
        int sepPos = fieldName.IndexOf(Constants.SUBFIELD_SEPARATOR);
        if (sepPos >= 0)
            {
            subFieldName = fieldName.Substring(sepPos + 1);
            fieldName = fieldName.Substring(0, sepPos);
            }
        return new FieldDef(fieldName, IsList, subFieldName);
        }
    }

public record class IndexedFieldDef(FieldDef FieldDef, int Index, int? SubIndex)
    {
    public static (IndexedFieldDef, string?) Create(FieldDef fieldDef, ExcelData excelData)
        {
        string? error = null;
        var fieldIndex = excelData.Headers.FindIndex(h => h == fieldDef.Name);
        int? subIndex = null;
        if (fieldIndex == -1)
            {
            error = ErrorDefs.FieldNotFound(fieldDef.Name);
            return (new IndexedFieldDef(fieldDef, -1, -1), error);
            }
        if (fieldDef.SubFieldName is not null)
            {
            var linkedData = excelData.LinkedData[fieldIndex]; //field "vestigingen" does not exist in linked data. Check this earlier.
            subIndex = linkedData.Headers.IndexOf(fieldDef.SubFieldName);
            if (subIndex == -1)
                {
                error = ErrorDefs.FieldNotFound(fieldDef.SubFieldName); //todo: make a more custom error for linked fields.
                return (new IndexedFieldDef(fieldDef, -1, -1), error);
                }
            }
        return (new IndexedFieldDef(fieldDef, fieldIndex, subIndex), null);
        }

    public static (IndexedFieldDef, string?) Create(Section section, ExcelData excelData)
        {
        FieldDef fieldDef = FieldDef.Parse(section.InbetweenRange.Text);
        return Create(fieldDef, excelData);
        }

    public string[] GetValues(List<string> row, ExcelData excelData)
        {
        if (this.SubIndex is null)
            return this.ToListIfNeeded(row[this.Index]);

        var keys = this.ToListIfNeeded(row[this.Index]);
        LinkedExcelData linkedData = excelData.LinkedData[this.Index];
        return keys.SelectMany(key =>
            {
                if(key != "")
                    {
                    var linkedRow = linkedData.GetRow(key); //todo: report error !
                    return this.ToListIfNeeded(linkedRow[(int)this.SubIndex]);
                    }
                return [];
            }).ToArray();
        }

    private string[] ToListIfNeeded(string value)
        {
        if(!this.FieldDef.IsList)
            return [value];
        return value.Split(';');
        }
    }