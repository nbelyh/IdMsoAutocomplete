using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using IdMsoAutocomplete.Configuration;
using Microsoft.VisualStudio.Language.Intellisense;
using OfficeOpenXml;

namespace IdMsoAutocomplete.DataLoaders
{
    public class Entry
    {
        public string ControlType { get; set; }
        public string TabSet { get; set; }
        public string Tab { get; set; }
        public string Group { get; set; }
        public string ParentControl { get; set; }

        public Completion Completion { get; set; }
    }

    public static class ExcelLoader
    {
        private static IDictionary<OfficeVersion, IDictionary<OfficeApplication, IEnumerable<Entry>>> _entries;

        public static IEnumerable<Entry> GetIdsFromCache(OfficeVersion officeVersion, OfficeApplication officeApplication)
        {
            if (_entries == null)
                _entries = new Dictionary<OfficeVersion, IDictionary<OfficeApplication, IEnumerable<Entry>>>();

            if (!_entries.ContainsKey(officeVersion))
                _entries[officeVersion] = new Dictionary<OfficeApplication, IEnumerable<Entry>>();

            var versionEntries = _entries[officeVersion];

            if (!versionEntries.ContainsKey(officeApplication))
                versionEntries[officeApplication] = LoadIds(officeVersion, officeApplication).OrderBy(e => e.Completion.DisplayText).ToList();

            return versionEntries[officeApplication];
        }

        private static IEnumerable<Entry> LoadIds(OfficeVersion officeVersion, OfficeApplication officeApplication)
        {
            var folder = Path.Combine(
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), 
                "OfficeData",
                officeVersion.ToString());

            var files = new DirectoryInfo(folder).GetFiles("*.xlsx")
                .Where(f => f.Name.StartsWith(officeApplication.ToString(), StringComparison.OrdinalIgnoreCase));

            foreach (var file in files)
            using (var package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets.First();
                for (var r = 2; r <= ws.Dimension.End.Row; r++)
                {
                    yield return new Entry
                    {
                        Completion = new Completion(ws.Cells[r, 1].GetValue<string>()),
                        ControlType = ws.Cells[r, 2].GetValue<string>(),
                        TabSet = ws.Cells[r, 3].GetValue<string>(),
                        Tab = ws.Cells[r, 4].GetValue<string>(),
                        Group = ws.Cells[r, 5].GetValue<string>(),
                        ParentControl = ws.Cells[r, 6].GetValue<string>(),
                    };
                }
            }
        }

        public static IEnumerable<Entry> GetIds(IServiceProvider serviceProvider)
        {
            var options = IdMsoPackage.GetOptions(serviceProvider);

            if (options.OfficeVersion == OfficeVersion.NotSpecified || options.OfficeApplication == OfficeApplication.NotSpecified)
            {
                OptionsPrompt.EditOptions(options);

                if (options.OfficeVersion == OfficeVersion.NotSpecified || options.OfficeApplication == OfficeApplication.NotSpecified)
                    return new List<Entry>();
            }

            return GetIdsFromCache(options.OfficeVersion, options.OfficeApplication).ToList();
        }
    }
}
