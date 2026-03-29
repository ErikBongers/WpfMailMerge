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

    }
