using System;
using Microsoft.VisualStudio.Shell;
using IdMsoAutocomplete.Configuration;
using Microsoft.VisualStudio.Shell.Interop;

namespace IdMsoAutocomplete
{
    using System.Runtime.InteropServices;

    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(PackageGuidString)]
    [ProvideOptionPage(typeof(Options), "IdMso Autocomplete", "General", 0, 0, true)]
    [ProvideProfileAttribute(typeof(Options), "IdMsoAutocomplete", "General", 0, 0, true)]
    public sealed class IdMsoPackage : Package
    {
        public const string PackageGuidString = "2bd31a92-3565-41c6-9c93-f0f112683544";

        private static Options _options;
        public static Options GetOptions(IServiceProvider serviceProvider)
        {
            if (_options == null)
            {
                var shell = serviceProvider.GetService(typeof(SVsShell)) as IVsShell;
                if (shell != null)
                {
                    IVsPackage package;
                    var packageToBeLoadedGuid = new Guid(PackageGuidString);
                    shell.LoadPackage(ref packageToBeLoadedGuid, out package);
                }
            }

            return _options;
        }

        protected override void Initialize()
        {
            base.Initialize();

            _options = GetDialogPage(typeof(Options)) as Options;
        }
    }
}
