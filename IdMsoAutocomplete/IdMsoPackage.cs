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
        private const string PackageGuidString = "2bd31a92-3565-41c6-9c93-f0f112683544";

        public static Options GetOptions(IServiceProvider serviceProvider)
        {
            var shell = serviceProvider.GetService(typeof(SVsShell)) as IVsShell;

            IVsPackage package;
            var packageToBeLoadedGuid = new Guid(PackageGuidString);
            shell.LoadPackage(ref packageToBeLoadedGuid, out package);

            var thisPackage = (IdMsoPackage) package;

            return thisPackage.GetDialogPage(typeof(Options)) as Options;
        }
    }
}
