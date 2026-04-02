using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfMailMerge;

public static class StringExtensions
    {
    public static string FirstWord(this string text)
        {
        int found = text.IndexOf(" ");
        return text.Substring(0, found == -1 ? text.Length : found);
        }

    }

public static class Tools
    {
    public static string? FindAbsolutePath(string relativePath, string basePath)
        {
        string generatedPath = relativePath;
        if (!File.Exists(generatedPath))
            {
            generatedPath = Path.Combine(basePath, relativePath);
            if (!File.Exists(generatedPath))
                return null;
            }
        return generatedPath;
        }

    public static (HashSet<string>, string) ExtractOptions(string markerText, IEnumerable<string> allOptions)
        {
        HashSet<string> options = [];
        string text = markerText;

        foreach (string optionName in allOptions)
            {
            int foundPos = text.IndexOf(" " + optionName); //todo: We're only checking a leading space. Not a trailing one or EOF.
            if (foundPos < 0)
                continue;
            options.Add(optionName);
            text = text.Replace(optionName, "");
            }

        return (options, text);
        }
    }
