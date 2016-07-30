using System.ComponentModel.Composition;
using Microsoft.VisualStudio.Utilities;
using Glyphfriend.Helpers;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Operations;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Editor;

namespace Glyphfriend.MsoImageCompletionProviders
{
    [ContentType("xml")]
    [Name("MsoImage")]
    [Export(typeof(ICompletionSourceProvider))]
    class MsoImageCompletionProvider : ICompletionSourceProvider
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
            return new MsoImageCompletionSource(
                textBuffer, 
                TextStructureNavigatorSelector.GetTextStructureNavigator(textBuffer), 
                ServiceProvider, 
                VsEditorAdaptersFactoryService);
        }
    }
}
