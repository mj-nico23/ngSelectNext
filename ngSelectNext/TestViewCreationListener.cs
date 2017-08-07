﻿using System;
using System.Collections.Generic;
using System.ComponentModel.Composition;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
        [Import]
        internal IEditorFormatMapService FormatMapService = null;

        [Import(typeof(IVsEditorAdaptersFactoryService))]
        internal IVsEditorAdaptersFactoryService editorFactory = null;

        public void VsTextViewCreated(IVsTextView textViewAdapter)
        {
            IWpfTextView textView = editorFactory.GetWpfTextView(textViewAdapter);

            if (textView != null)
                AddCommandFilter(textViewAdapter, textView, new MultiPointEditCommandFilter(textView));
        }

        private void AddCommandFilter(IVsTextView viewAdapter, IWpfTextView textView, MultiPointEditCommandFilter commandFilter)
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