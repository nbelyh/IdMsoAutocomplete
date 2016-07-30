using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using IdMsoAutocomplete;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Shell;
using Task = System.Threading.Tasks.Task;

namespace Glyphfriend.Helpers
{
    public class Entry
    {
        public string ControlName { get; set; }
        public string ControlType { get; set; }
        public string TabSet { get; set; }
        public string Tab { get; set; }
        public string Group { get; set; }
        public string ParentControl { get; set; }
    }

    public static class ExcelLoader
    {
        public static IEnumerable<Entry> Entries { get; private set; }

        static IEnumerable<Entry> LoadFile(OfficeApplication app, OfficeVersion officeVersion)
        {
            var folder = Path.Combine(
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), 
                "OfficeData",
                officeVersion.ToString());

            var files = new DirectoryInfo(folder).GetFiles("*.xlsx")
                .Where(f => f.Name.StartsWith(app.ToString(), StringComparison.OrdinalIgnoreCase));

            foreach (var file in files)
            using (var package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets.First();
                for (var r = 2; r <= ws.Dimension.End.Row; r++)
                {
                    yield return new Entry
                    {
                        ControlName = ws.Cells[r, 1].GetValue<string>(),
                        ControlType = ws.Cells[r, 2].GetValue<string>(),
                        TabSet = ws.Cells[r, 3].GetValue<string>(),
                        Tab = ws.Cells[r, 4].GetValue<string>(),
                        Group = ws.Cells[r, 5].GetValue<string>(),
                        ParentControl = ws.Cells[r, 6].GetValue<string>(),
                    };
                }
            }
        }

        public static IEnumerable<Completion> ToIdList(this IEnumerable<Entry> input)
        {
            return input
                .Select(e => e.ControlName)
                .Distinct()
                .Select(e => new Completion(e));
        }

        public static void LoadEntries()
        {
            Entries = LoadFile(OfficeApplication.Visio, OfficeVersion.Office2010).ToList();
        }
    }
}
