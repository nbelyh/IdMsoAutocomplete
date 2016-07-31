using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Utilities;

namespace IdMsoAutocomplete.Completion
{
    [ContentType("xml")]
    [Name("MsoImage")]
    [Export(typeof(ICompletionSourceProvider))]
    class CompletionSourceProvider : ICompletionSourceProvider
    {
        [Import(typeof(SVsServiceProvider))]
        internal System.IServiceProvider ServiceProvider { get; set; }

        [Import]
        internal IVsEditorAdaptersFactoryService VsEditorAdaptersFactoryService { get; set; }

        public ICompletionSource TryCreateCompletionSource(ITextBuffer textBuffer)
        {
            // Build the MsoImage completion handler for the current text buffer
            return new CompletionSource(
                textBuffer, 
                ServiceProvider, 
                VsEditorAdaptersFactoryService);
        }
    }
}
