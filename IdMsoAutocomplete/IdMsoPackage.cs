using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.IO.Compression;
using System.Reflection;
using Microsoft.VisualStudio;

namespace IdMsoAutocomplete
{
    using System.ComponentModel;
    using System.Runtime.InteropServices;
    using ImageList = Dictionary<string, ImageSource>;

    public enum OfficeVersion
    {
        Office2010,
        Office2013,
        Office2016,
    }

    public enum OfficeApplication
    {
        Access,
        Excel,
        InfoPath,
        OneNote,
        Outlook,
        PowerPoint,
        Project,
        Publisher,
        SharePointDesigner,
        Visio,
        Word
    }

    public class IdMsoOptions : DialogPage
    {
        [Category("Office Application")]
        [DisplayName("Office Application")]
        public OfficeApplication OfficeApplication
        {
            get;
            set;
        }

        [Category("Office Application")]
        [DisplayName("Application Version")]
        public OfficeVersion OfficeVersion
        {
            get;
            set;
        }

        public override void LoadSettingsFromStorage()
        {
            OfficeVersion = OfficeVersion.Office2010;
            OfficeApplication = OfficeApplication.Visio;
            base.LoadSettingsFromStorage();
        }
    }

    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(PackageGuidString)]
    [ProvideOptionPage(typeof(IdMsoOptions), "IdMso Autocomplete", "General", 0, 0, true)]
    public sealed class IdMsoPackage : Package
    {
        public const string PackageGuidString = "2bd31a92-3565-41c6-9c93-f0f112683544";

        private static Dictionary<OfficeVersion, ImageList> _msoImages;

        public static ImageList GetMsoImages(OfficeVersion officeVersion)
        {
            LoadMsoImages(officeVersion);
            return _msoImages[officeVersion];
        }

        public static void LoadMsoImages(OfficeVersion officeVersion)
        {
            if (_msoImages == null)
                _msoImages = new Dictionary<OfficeVersion, ImageList>();

            if (!_msoImages.ContainsKey(officeVersion))
                _msoImages[officeVersion] = LoadGlyphsFromFile(officeVersion);
        }

        private static ImageList LoadGlyphsFromFile(OfficeVersion officeVersion)
        {
            // Instantiate the dictionary with the name of the library
            var result = new ImageList();

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
                    using (var glyphStream = glyph.Open())
                    using (var ms = new MemoryStream())
                    {
                        glyphStream.CopyTo(ms);

                        try
                        {
                            // Generate the Glyph
                            var bitmap = new BitmapImage();
                            bitmap.BeginInit();
                            bitmap.StreamSource = ms;
                            bitmap.CacheOption = BitmapCacheOption.OnLoad;
                            bitmap.EndInit();
                            bitmap.Freeze();

                            // Add it to the collection (under the specific library)
                            result.Add(Path.GetFileNameWithoutExtension(glyph.Name), bitmap);
                        }
                        catch (Exception)
                        {
                            // An error occurred when creating the glyph, ignore it (it will be handled in the provider)
                        }
                    }
                }
            }

            return result;
        }
    }
}
