using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Language.Intellisense;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.XmlEditor;

namespace IdMsoAutocomplete.CompletionSupport
{
    class CompletionSource : ICompletionSource
    {
        private readonly ITextBuffer _buffer;
        private IEnumerable<Entry> _entries;

        private readonly IVsEditorAdaptersFactoryService _vsEditorAdaptersFactoryService;
        private readonly IServiceProvider _serviceProvider;
        private readonly XmlLanguageService _xmlLanguageService;
        private bool _disposed;

        public CompletionSource(
            ITextBuffer buffer,
            IServiceProvider serviceProvider,
            IVsEditorAdaptersFactoryService vsEditorAdaptersFactoryService)
        {
            _buffer = buffer;
            _serviceProvider = serviceProvider;
            _vsEditorAdaptersFactoryService = vsEditorAdaptersFactoryService;

            _xmlLanguageService = (XmlLanguageService)_serviceProvider.GetService(typeof(XmlLanguageService));
        }

        private IEnumerable<Entry> GetEntries(string type)
        {
            return _entries.Where(e => e.ControlType == type);
        }

        static readonly string[] SupportedAttributeNames =
        {
            "idMso",
            "insertBeforeMso",
            "insertAfterMso"
        };

        private static readonly string[] SupportedControlsTypes =
        {
            "separator",
            "box",
            "button",
            "buttonGroup",
            "editBox",
            "toggleButton",
            "splitButton",
            "menu",
            "dynamicMenu",
            "menuSeparator",
            "separator",
            "gallery",
            "checkBox",
            "comboBox",
            "control",
            "labelControl",
        };

        public static readonly string[] SupportedNamespaces =
        {
            "http://schemas.microsoft.com/office/2009/07/customui",
            "http://schemas.microsoft.com/office/2006/01/customui",
        };

        public void AugmentCompletionSession(ICompletionSession session, IList<CompletionSet> completionSets)
        {
            // If this session has already been disposed, ignore it
            if (_disposed)
                return;

            IVsTextView vsTextView = _vsEditorAdaptersFactoryService.GetViewAdapter(session.TextView);

            int line;
            int column;
            vsTextView.GetCaretPos(out line, out column);

            var doc = _xmlLanguageService.GetParseTree(
                _xmlLanguageService.GetSource(vsTextView), vsTextView, line, column, Microsoft.VisualStudio.Package.ParseReason.CompleteWord);

            if (doc == null || !SupportedNamespaces.Contains(doc.RootNamespaceURI))
                return;

            var nf = new NodeFinder(line, column);
            nf.Visit(doc);

            if (nf.Scope is XmlAttribute)
            {
                var attr = (XmlAttribute)nf.Scope;

                var attrName = attr.LocalName;

                if (SupportedAttributeNames.Contains(attrName))
                {
                    if (_entries == null)
                        _entries = ExcelLoader.GetIds(_serviceProvider);

                    var elt = (XmlElement)attr.Parent;
                    IEnumerable<Entry> completions = null;

                    if (elt == null)
                        return;

                    var type = elt.LocalName;

                    if (string.IsNullOrEmpty(type))
                        return;

                    if (type == "tab")
                    {
                        completions = GetEntries(type);

                        var eltParent = elt.Parent as XmlElement;
                        if (eltParent != null && eltParent.LocalName == "backstage")
                            completions = completions.Where(e => e.TabSet == "None (Backstage View)");
                        else if (eltParent != null && eltParent.LocalName == "tabSet")
                            completions = completions.Where(e => e.TabSet != "None (Core Tab)");
                        else
                            completions = completions.Where(e => e.TabSet == "None (Core Tab)");
                    }
                    else if (type == "tabSet")
                    {
                        completions = GetEntries(type);
                    }
                    else if (type == "group")
                    {
                        completions = GetEntries(type);

                        var tabName = GetParentItemIdMso(elt, "tab");

                        if (!string.IsNullOrEmpty(tabName))
                            completions = completions.Where(e => e.Tab == tabName);
                    }
                    else if (SupportedControlsTypes.Contains(type))
                    {
                        var tab = GetParentItemIdMso(elt, "tab");
                        var group = GetParentItemIdMso(elt, "group");
                        var contextMenu = GetParentItemIdMso(elt, "contextMenu");

                        if (attrName == "idMso")
                            completions = GetEntries(type);
                        else
                            completions = _entries.Where(e => SupportedControlsTypes.Contains(e.ControlType));

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
                    }
                    else if (type == "command")
                    {
                        completionSets.Clear();

                        completions = _entries
                            .Where(e => SupportedControlsTypes.Contains(e.ControlType));
                    }
                    else if (type == "category")
                    {
                        completions = GetEntries(type);

                        var tabName = GetParentItemIdMso(elt, "tab");

                        if (!string.IsNullOrEmpty(tabName))
                            completions = completions.Where(e => e.Tab == tabName);

                        var taskGroupName = GetParentItemIdMso(elt, "taskGroup");

                        if (!string.IsNullOrEmpty(tabName))
                            completions = completions.Where(e => e.Group == taskGroupName);
                    }
                    else if (type == "taskFormGroup" || type == "taskGroup")
                    {
                        completions = GetEntries(type);

                        var tabName = GetParentItemIdMso(elt, "tab");

                        if (!string.IsNullOrEmpty(tabName))
                            completions = completions.Where(e => e.Tab == tabName);
                    }
                    else if (type == "task")
                    {
                        completions = GetEntries(type);

                        var taskGroup = GetParentItemIdMso(elt, "taskGroup");
                        if (!string.IsNullOrEmpty(taskGroup))
                            completions = completions.Where(e => e.Group == taskGroup);

                        var category = GetParentItemIdMso(elt, "category");
                        if (!string.IsNullOrEmpty(taskGroup))
                            completions = completions.Where(e => e.ParentControl == category);
                    }
                    else if (type == "contextMenu")
                    {
                        completions = GetEntries(type);
                    }

                    if (completions != null)
                    {
                        completionSets.Clear();
                        completionSets.Add(new CompletionSet("idMso", "idMso", CreateTrackingSpan(session),
                            completions.Select(e => e.Completion), null));
                    }
                }
                else if (attr.LocalName == "imageMso")
                {
                    var completions = ImageLoader.GetMsoImages(_serviceProvider);

                    completionSets.Clear();
                    completionSets.Add(new CompletionSet("imageMso", "imageMso", CreateTrackingSpan(session), completions, null));
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

    }
}
