using System.ComponentModel.Composition;
using System.Linq;
using IdMsoAutocomplete.Configuration;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;
using Microsoft.XmlEditor;
using ErrorHandler = Microsoft.VisualStudio.ErrorHandler;
using Task = System.Threading.Tasks.Task;

namespace IdMsoAutocomplete.Completion
{
    [Export(typeof(IVsTextViewCreationListener))]
    [ContentType("xml")]
    [TextViewRole(PredefinedTextViewRoles.Document)]
    class TextViewListener : IVsTextViewCreationListener
    {
        [Import]
        internal IVsEditorAdaptersFactoryService AdaptersFactory { get; set; }

        [Import]
        internal ICompletionBroker CompletionBroker { get; set; }

        [Import(typeof(SVsServiceProvider))]
        internal System.IServiceProvider ServiceProvider { get; set; }

        // This will be triggered when a xml document is opened
        public void VsTextViewCreated(IVsTextView textViewAdapter)
        {
            // Set up the Completion handler for xml documents
            var view = AdaptersFactory.GetWpfTextView(textViewAdapter);

            var xmlLanguageService = (XmlLanguageService)ServiceProvider.GetService(typeof(XmlLanguageService));
            var quoteChar = xmlLanguageService.XmlPrefs.AutoInsertAttributeQuotes;

            if (XmlFileIsRibbon(view))
                PreLoadData();

            var filter = new CommandFilter(view, CompletionBroker, quoteChar);

            IOleCommandTarget next;
            ErrorHandler.ThrowOnFailure(textViewAdapter.AddCommandFilter(filter, out next));
            filter.Next = next;
        }

        private void PreLoadData()
        {
            var options = IdMsoPackage.GetOptions(ServiceProvider);

            if (options.OfficeApplication != OfficeApplication.NotSpecified &&
                options.OfficeVersion != OfficeVersion.NotSpecified)
                Task.Run(() => ExcelLoader.GetIdsFromCache(options.OfficeVersion, options.OfficeApplication));

            if (options.OfficeApplication != OfficeApplication.NotSpecified &&
                options.OfficeVersion != OfficeVersion.NotSpecified)
                Task.Run(() => ImageLoader.GetMsoImagesFromCache(options.OfficeVersion));
        }

        private static bool XmlFileIsRibbon(IWpfTextView view)
        {
            var snapshot = view.TextBuffer.CurrentSnapshot;
            var length = snapshot.LineCount < 4000 ? snapshot.LineCount : 4000;

            var textBegin = snapshot.GetText(0, length);
            return CompletionSource.SupportedNamespaces.Any(ns => textBegin.Contains(ns));
        }
    }
}
