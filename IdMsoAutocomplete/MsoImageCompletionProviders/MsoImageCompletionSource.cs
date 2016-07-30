using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Operations;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Media;
using Glyphfriend.Helpers;
using Microsoft.XmlEditor;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using System.Threading.Tasks;
using IdMsoAutocomplete;
using Microsoft.VisualStudio.Shell.Interop;

namespace Glyphfriend.MsoImageCompletionProviders
{
    class MsoImageCompletionSource : ICompletionSource
    {
        private readonly ITextBuffer _buffer;
        private List<Completion> _images;
        private IEnumerable<Entry> _entries;

        private IVsEditorAdaptersFactoryService _vsEditorAdaptersFactoryService;
        private ITextStructureNavigator _textStructureNavigator;
        private IServiceProvider _serviceProvider;
        private bool _disposed = false;

        public MsoImageCompletionSource(
            ITextBuffer buffer,
            ITextStructureNavigator textStructureNavigator,
            IServiceProvider serviceProvider,
            IVsEditorAdaptersFactoryService vsEditorAdaptersFactoryService)
        {
            _buffer = buffer;
            _textStructureNavigator = textStructureNavigator;
            _serviceProvider = serviceProvider;
            _vsEditorAdaptersFactoryService = vsEditorAdaptersFactoryService;
        }

        public IEnumerable<Entry> GetEntries(string type)
        {
            return _entries.Where(e => e.ControlType == type);
        }

        static readonly string[] idMso =
        {
            "idMso",
            "insertBeforeMso",
            "insertAfterMso"
        };

        public void AugmentCompletionSession(ICompletionSession session, IList<CompletionSet> completionSets)
        {
            // If this session has already been disposed, ignore it
            if (_disposed)
                return;

            XmlLanguageService xls = (XmlLanguageService)_serviceProvider.GetService(typeof(XmlLanguageService));

            IVsTextView vsTextView = _vsEditorAdaptersFactoryService.GetViewAdapter(session.TextView);

            int line;
            int column;
            int tmp = vsTextView.GetCaretPos(out line, out column);

            XmlDocument doc = xls.GetParseTree(xls.GetSource(vsTextView), vsTextView, line, column, Microsoft.VisualStudio.Package.ParseReason.CompleteWord);

            NodeFinder nf = new NodeFinder(line, column);
            nf.Visit(doc);

            if (nf.Scope is XmlAttribute)
            {
                XmlAttribute attr = (XmlAttribute)nf.Scope;


                if (idMso.Contains(attr.LocalName))
                {
                    if (_entries == null)
                    {
                        ExcelLoader.LoadEntries();
                        _entries = ExcelLoader.Entries;
                    }

                    XmlElement elt = (XmlElement)attr.Parent;
                    IEnumerable<Entry> completions = null;

                    if (elt == null)
                        return;

                    var type = elt.LocalName;

                    if (string.IsNullOrEmpty(type))
                        return;

                    switch (type)
                    {
                        case "tab":
                            {
                                completions = GetEntries(type);

                                var eltParent = elt.Parent as XmlElement;
                                bool backstage = (eltParent != null && eltParent.LocalName == "backstage");

                                completions = completions.Where(e => (e.TabSet == "None (Backstage View)") == backstage);

                                break;
                            }

                        case "group":
                            {
                                completions = GetEntries(type);

                                var tabName = GetParentItemIdMso(elt, "tab");

                                if (!string.IsNullOrEmpty(tabName))
                                    completions = completions.Where(e => e.Tab == tabName);
                                break;
                            }

                        case "button":
                        case "toggleButton":
                        case "splitButton":
                        case "menu":
                        case "gallery":
                        case "checkBox":
                        case "comboBox":
                        case "control":
                        case "labelControl":
                            {
                                var tab = GetParentItemIdMso(elt, "tab");
                                var group = GetParentItemIdMso(elt, "group");
                                var contextMenu = GetParentItemIdMso(elt, "contextMenu");

                                completions = GetEntries(type);

                                if (tab == "None (Context Menu)")
                                {
                                    if (!string.IsNullOrEmpty(contextMenu))
                                        completions = completions.Where(e => e.Group == contextMenu);
                                }
                                else if (tab != "None (Backstage View)")
                                {
                                    if (!string.IsNullOrEmpty(tab))
                                        completions = completions.Where(e => e.Tab == tab);

                                    if (!string.IsNullOrEmpty(group))
                                        completions = completions.Where(e => e.Group == group);
                                }

                                break;
                            }

                        case "command":
                            {
                                completionSets.Clear();

                                completions = _entries
                                    .Where(e => e.ControlType != "contextMenu");

                                break;
                            }

                        case "category":
                            {
                                completions = GetEntries(type);

                                var tabName = GetParentItemIdMso(elt, "tab");

                                if (!string.IsNullOrEmpty(tabName))
                                    completions = completions.Where(e => e.Tab == tabName);

                                var taskGroupName = GetParentItemIdMso(elt, "taskGroup");

                                if (!string.IsNullOrEmpty(tabName))
                                    completions = completions.Where(e => e.Group == taskGroupName);

                                break;
                            }

                        case "taskFormGroup":
                        case "taskGroup":
                            {
                                completions = GetEntries(type);

                                var tabName = GetParentItemIdMso(elt, "tab");

                                if (!string.IsNullOrEmpty(tabName))
                                    completions = completions.Where(e => e.Tab == tabName);

                                break;
                            }

                        case "task":
                            {
                                completions = GetEntries(type);

                                var taskGroup = GetParentItemIdMso(elt, "taskGroup");
                                if (!string.IsNullOrEmpty(taskGroup))
                                    completions = completions.Where(e => e.Group == taskGroup);

                                var category = GetParentItemIdMso(elt, "category");
                                if (!string.IsNullOrEmpty(taskGroup))
                                    completions = completions.Where(e => e.ParentControl == category);

                                break;
                            }

                        case "contextMenu":
                            {
                                completions = GetEntries(type);

                                break;
                            }
                    }

                    if (completions != null)
                    {
                        completionSets.Clear();
                        completionSets.Add(new CompletionSet("idMso", "idMso", CreateTrackingSpan(session), completions.ToIdList(), null));
                    }
                }
                else if (attr.LocalName == "imageMso")
                {
                    if (_images == null)
                    {
                        IdMsoPackage.LoadMsoImages(OfficeVersion.Office2010);
                        // Build a collection of MsoImages to handle 
                        _images = IdMsoPackage.GetMsoImages(OfficeVersion.Office2010)
                            .Select(e => MsoImageCompletion(e.Key, e.Value))
                            .ToList();
                    }

                    completionSets.Clear();
                    completionSets.Add(new CompletionSet("imageMso", "imageMso", CreateTrackingSpan(session), _images.OrderBy(c => c.DisplayText), null));
                }
            }
        }

        private static string GetParentItemIdMso(XmlElement elt, string type)
        {
            var eltParent = elt.Parent as XmlElement;
            if (eltParent == null)
                return null;

            if (eltParent.LocalName != type)
                return GetParentItemIdMso(eltParent, type);

            return eltParent.GetAttribute("idMso");
        }

        private ITrackingSpan CreateTrackingSpan(ICompletionSession session)
        {
            int position = session.GetTriggerPoint(session.TextView.TextBuffer).GetPosition(_buffer.CurrentSnapshot);

            char[] separators = new[] { '"', '\'', '.' };

            string text = _buffer.CurrentSnapshot.GetText();
            int first = text.Substring(0, position).LastIndexOfAny(separators) + 1;

            return _buffer.CurrentSnapshot.CreateTrackingSpan(new Span(first, position - first), SpanTrackingMode.EdgeInclusive);
        }

        public void Dispose()
        {
            _disposed = true;
        }

        private static Completion MsoImageCompletion(string msoImage, ImageSource MsoImageImage)
        {
            // Map a completion object for each MsoImage to the appropriate image
            var formattedMsoImage = msoImage;
            // Build a completion for each MsoImage
            return new Completion(formattedMsoImage, formattedMsoImage, formattedMsoImage, MsoImageImage, formattedMsoImage);
        }

        private ITrackingSpan FindTokenSpanAtPosition(ICompletionSession session)
        {
            // Look for the nearly start of line or space prior to the starting character and 
            // examine it
            SnapshotPoint currentPoint = session.TextView.Caret.Position.BufferPosition - 1;
            TextExtent extent = _textStructureNavigator.GetExtentOfWord(currentPoint);
            return currentPoint.Snapshot.CreateTrackingSpan(extent.Span, SpanTrackingMode.EdgeInclusive);
        }


    }
}
