using System.ComponentModel.Composition;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Text.Classification;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Utilities;

namespace ngSelectNext
{
    [Export(typeof(IVsTextViewCreationListener))]
    [ContentType("text")]
    [TextViewRole(PredefinedTextViewRoles.Document)]
    internal class TestViewCreationListener : IVsTextViewCreationListener
    {
        [Export(typeof(AdornmentLayerDefinition))]
        [Name("SelectNextLayer")]
        [TextViewRole(PredefinedTextViewRoles.Editable)]
        internal AdornmentLayerDefinition multiEditAdornmentLayer;

        [Import]
        internal IEditorFormatMapService FormatMapService;

        [Import(typeof(IVsEditorAdaptersFactoryService))]
        internal IVsEditorAdaptersFactoryService editorFactory;

        public void VsTextViewCreated(IVsTextView textViewAdapter)
        {
            IWpfTextView textView = editorFactory.GetWpfTextView(textViewAdapter);

            if (textView != null)
                AddCommandFilter(textViewAdapter, textView, new SelectNextCommandFilter(textView));
        }

        private void AddCommandFilter(IVsTextView viewAdapter, IWpfTextView textView, SelectNextCommandFilter commandFilter)
        {
            if (commandFilter.Added == false)
            {
                IOleCommandTarget next;
                int result = viewAdapter.AddCommandFilter(commandFilter, out next);

                if (result == VSConstants.S_OK)
                {
                    commandFilter.Added = true;
                    textView.Properties.AddProperty(typeof(SelectNextCommand), commandFilter);

                    if (next != null)
                    {
                        commandFilter.NextTarget = next;
                    }
                }
            }
        }
    }
}
