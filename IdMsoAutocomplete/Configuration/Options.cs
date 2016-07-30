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
    public class Options : DialogPage, INotifyPropertyChanged
    {
        private OfficeApplication _officeApplication;

        [Category("Target Office Version")]
        [DisplayName("Office Application")]
        [DefaultValue(OfficeApplication.NotSpecified)]
        public OfficeApplication OfficeApplication
        {
            get
            {
                return _officeApplication;
            }
            set
            {
                _officeApplication = value;
                OnPropertyChange("OfficeApplication");
            }
        }

        private OfficeVersion _officeVersion;

        [Category("Target Office Version")]
        [DisplayName("Office Version")]
        [DefaultValue(OfficeVersion.NotSpecified)]
        public OfficeVersion OfficeVersion
        {
            get
            {
                return _officeVersion;
            }
            set
            {
                _officeVersion = value;
                OnPropertyChange("OfficeVersion");
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        private void OnPropertyChange(string property)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(property));
            }
        }
    }

}