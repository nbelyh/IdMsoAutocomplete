using Microsoft.VisualStudio.Shell;
using IdMsoAutocomplete.Configuration;

namespace IdMsoAutocomplete
{
    using System.Runtime.InteropServices;

    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(IdMsoPackage.PackageGuidString)]
    [ProvideOptionPage(typeof(Options), "IdMso Autocomplete", "General", 0, 0, true)]
    [ProvideProfileAttribute(typeof(Options), "IdMsoAutocomplete", "General", 0, 0, true)]
    public sealed class IdMsoPackage : Package
    {
        public const string PackageGuidString = "2bd31a92-3565-41c6-9c93-f0f112683544";

        public static Options Options { get; private set; }

        protected override void Initialize()
        {
            base.Initialize();

            Options = GetDialogPage(typeof(Options)) as Options;
        }
    }
}
