using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Utilities;

namespace IdMsoAutocomplete.CompletionProviders
{
    [ContentType("xml")]
    [Name("MsoImage")]
    [Export(typeof(ICompletionSourceProvider))]
    class CompletionSourceProvider : ICompletionSourceProvider
    {
        [Import]
        public ITextStructureNavigatorSelectorService TextStructureNavigatorSelector = null;

        [Import(typeof(SVsServiceProvider))]
        internal SVsServiceProvider ServiceProvider { get; set; }

        [Import]
        internal IVsEditorAdaptersFactoryService VsEditorAdaptersFactoryService { get; set; }

        public ICompletionSource TryCreateCompletionSource(ITextBuffer textBuffer)
        {
            // Build the MsoImage completion handler for the current text buffer
            return new CompletionSource(
                textBuffer, 
                TextStructureNavigatorSelector.GetTextStructureNavigator(textBuffer), 
                ServiceProvider, 
                VsEditorAdaptersFactoryService);
        }
    }
}
