using System.ComponentModel;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace IdMsoAutocomplete.Configuration
{
    public enum OfficeApplication
    {
        NotSpecified,
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

    public enum OfficeVersion
    {
        NotSpecified,
        Office2010,
        Office2013,
        Office2016,
    }

    [ClassInterface(ClassInterfaceType.AutoDual)]
    [Guid("D9183604-3546-4151-848F-66EEB4EDB2AE")]
    public class Options : DialogPage
    {
        [Category("Target Office Version")]
        [DisplayName("Office Application")]
        [DefaultValue(OfficeApplication.NotSpecified)]
        public OfficeApplication OfficeApplication
        {
            get;
            set;
        }

        [Category("Target Office Version")]
        [DisplayName("Office Version")]
        [DefaultValue(OfficeVersion.NotSpecified)]
        public OfficeVersion OfficeVersion
        {
            get;
            set;
        }
    }

}