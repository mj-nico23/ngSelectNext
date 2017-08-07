//------------------------------------------------------------------------------
// <copyright file="SelectNextCommand.cs" company="Company">
//     Copyright (c) Company.  All rights reserved.
// </copyright>
//------------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.Diagnostics;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;

namespace ngSelectNext
{
    internal sealed class SelectNextCommand
    {
        public const int CommandId = 0x0100;

        public static readonly Guid CommandSet = new Guid("8cfe5bed-9dd1-4111-bdfb-708bef268c4f");

        private readonly Package package;

        private SelectNextCommand(Package package)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));

            OleMenuCommandService commandService = ServiceProvider.GetService(typeof(IMenuCommandService)) as OleMenuCommandService;
            if (commandService != null)
            {
                var menuCommandID = new CommandID(CommandSet, CommandId);
                var menuItem = new MenuCommand(MenuItemCallback, menuCommandID);
                commandService.AddCommand(menuItem);

            }
        }

        public static SelectNextCommand Instance
        {
            get;
            private set;
        }

        private IServiceProvider ServiceProvider
        {
            get
            {
                return package;
            }
        }

        public static void Initialize(Package package)
        {

            Instance = new SelectNextCommand(package);
        }

        private void MenuItemCallback(object sender, EventArgs e)
        {
            IWpfTextView textview = Helpers.GetCurentTextView();

            if (textview == null)
            {
                Debug.WriteLine("Could not find IWpfTextView");
                return;
            }

            ITextSnapshot snapshot = textview.TextSnapshot;

            if (snapshot != snapshot.TextBuffer.CurrentSnapshot)
                return;

            if (textview.Selection.IsEmpty)
            {
                return;
            }

            if (MultiPointEditCommandFilter.m_trackList == null)
                MultiPointEditCommandFilter.m_trackList = new List<ITrackingPoint>();

            ITextSelection selection = textview.Selection;

            //MultiPointEditCommandFilter.SelectionsMade.Add(selection);

            string currentSelection = selection.SelectedSpans[textview.Selection.SelectedSpans.Count - 1]
                .GetText();

            //Find next
            var intNext = snapshot.GetText()
                .IndexOf(
                    currentSelection,
                    textview.Selection.Start.Position + currentSelection.Length - 1);

            if (intNext > -1)
            {
                var curPosition = textview.Caret.Position;

                var curTrackPoint = textview.TextSnapshot.CreateTrackingPoint(curPosition.BufferPosition.Position, PointTrackingMode.Positive);

                MultiPointEditCommandFilter.m_trackList.Add(curTrackPoint);

                SnapshotSpan sp = new SnapshotSpan(snapshot, intNext, currentSelection.Length);
                textview.Selection.Select(sp, false);
            }

            var initActive = textview.Selection.ActivePoint;
            //var newAnchor = initAnchor.TranslateTo(textview.TextSnapshot, PointTrackingMode.Negative);
            var newActive = initActive.TranslateTo(textview.TextSnapshot, PointTrackingMode.Negative);
            //textview.Selection.Select(newAnchor, newActive);
            textview.Caret.MoveTo(newActive, PositionAffinity.Predecessor);
        }


    }
}
