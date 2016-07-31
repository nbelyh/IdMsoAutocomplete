using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Windows.Media.Imaging;
using IdMsoAutocomplete.Configuration;
using Microsoft.VisualStudio.Language.Intellisense;

namespace IdMsoAutocomplete.DataLoaders
{
    public static class ImageLoader
    {
        private static Dictionary<OfficeVersion, IEnumerable<Completion>> _msoImages;

        public static IEnumerable<Completion> GetMsoImages(IServiceProvider serviceProvider)
        {
            var options = IdMsoPackage.GetOptions(serviceProvider);

            if (options.OfficeVersion == OfficeVersion.NotSpecified)
            {
                OptionsPrompt.EditOptions(options);

                if (options.OfficeVersion == OfficeVersion.NotSpecified)
                    return new List<Completion>();
            }

            return GetMsoImagesFromCache(options.OfficeVersion);
        }

        public static IEnumerable<Completion> GetMsoImagesFromCache(OfficeVersion officeVersion)
        {
            if (_msoImages == null)
                _msoImages = new Dictionary<OfficeVersion, IEnumerable<Completion>>();

            if (!_msoImages.ContainsKey(officeVersion))
                _msoImages[officeVersion] = LoadGlyphsFromFile(officeVersion).OrderBy(e => e.DisplayText).ToList();

            return _msoImages[officeVersion];
        }

        private static IEnumerable<Completion> LoadGlyphsFromFile(OfficeVersion officeVersion)
        {
            var folder = Path.Combine(
                Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                "OfficeData",
                officeVersion.ToString());

            var files = new DirectoryInfo(folder).GetFiles("*.zip");

            foreach (var file in files)
                using (var stream = File.OpenRead(file.FullName))
                using (var zip = new ZipArchive(stream))
                {
                    // Iterate through all of the glyphs in this dictionary
                    foreach (var glyph in zip.Entries)
                    {
                        var name = Path.GetFileNameWithoutExtension(glyph.Name);

                        using (var glyphStream = glyph.Open())
                        using (var ms = new MemoryStream())
                        {
                            glyphStream.CopyTo(ms);

                            // Generate the Glyph
                            var bitmap = new BitmapImage();
                            bitmap.BeginInit();
                            bitmap.StreamSource = ms;
                            bitmap.CacheOption = BitmapCacheOption.OnLoad;
                            bitmap.EndInit();
                            bitmap.Freeze();

                            // Add it to the collection (under the specific library)
                            yield return new Completion(name, name, name, bitmap, name);
                        }
                    }
                }
        }
    }
}