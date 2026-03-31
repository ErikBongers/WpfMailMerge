using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WpfMailMerge
    {
    internal static class ErrorDefs
        {
        public static string FieldNotFound(string fieldName) { return $"Field \"{fieldName}\" not found"; }
        public static string OnlyFieldsWithoutTextExpected (string markerText) { return $"Found text mixed with fields in field-only marker: \"{markerText}\""; }
        public static string UnknownMarker (string marker) { return $"Unknown marker: \"%%{marker}\""; }
        public static string CanNotOpenFile(string fileName) { return $"Cannot open file \"{fileName}\""; }
        public static string MoreThanOneMarker(string marker) { return $"Found more than one \"{marker}\" marker."; }
        public static string MissingMarker(string marker) { return $"Missing \"{marker}\" marker."; }
        public static string MissingEndOfPlaceHolder(string marker) { return $"Missing \"{marker}\" to close placeholder."; }
        public static string EndSectionWithoutBeginning(string marker) { return $"Found end section marker for \"{marker}\" but no begin."; }
        public static string SectionWithoutEndMarker(string marker) { return $"No end marker for section \"{marker}\""; }
        public static string FileToIncludeNotFound(string path) { return $"File to include not found: {path}"; }
        public static string TooManyLinkFields(string fileName, string rangeName) { return $"Too many link fields in workbook {fileName} range {rangeName}."; }
        }
    }
