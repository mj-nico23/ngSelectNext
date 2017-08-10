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
        public static SelectNextTextAdornment adornment;
        public const int CommandId = 0x0100;
        public const int SkipCommandId = 0x0105;

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

                var menuSkipCommandID = new CommandID(CommandSet, SkipCommandId);
                var menuSkipItem = new MenuCommand(MenuItemCallback, menuSkipCommandID);
                commandService.AddCommand(menuSkipItem);

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
            bool skip_next = false;

            MenuCommand obj_sender = (MenuCommand)sender;

            if (obj_sender.CommandID.ID == SkipCommandId)
                skip_next = true;

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

            if (SelectNextCommandFilter.m_trackList == null)
                SelectNextCommandFilter.m_trackList = new List<ITrackingPoint>();

            ITextSelection selection = textview.Selection;

            PointTrackingMode trackingMode = selection.IsReversed ? PointTrackingMode.Negative : PointTrackingMode.Positive;

            //MultiPointEditCommandFilter.SelectionsMade.Add(selection);

            string currentSelection = selection.SelectedSpans[textview.Selection.SelectedSpans.Count - 1]
                .GetText();

            //Find next
            var intNext = snapshot.GetText()
                .IndexOf(
                    currentSelection,
                    textview.Selection.Start.Position.Position + currentSelection.Length, StringComparison.Ordinal);

            if (intNext == -1)
            {
                intNext = snapshot.GetText().IndexOf(currentSelection, 0, StringComparison.Ordinal);
            }

            if (intNext > -1)
            {
                if (!skip_next)
                {
                    var curPosition = textview.Caret.Position;

                    var curTrackPoint = textview.TextSnapshot.CreateTrackingPoint(curPosition.BufferPosition.Position,
                        trackingMode);

                    SelectNextCommandFilter.m_trackList.RemoveAll(
                        point => point.GetPosition(snapshot) == curTrackPoint.GetPosition(snapshot));

                    SelectNextCommandFilter.m_trackList.Add(curTrackPoint);

                    if (adornment == null)
                        adornment = new SelectNextTextAdornment(textview);

                    adornment.CreateVisuals();
                }

                SnapshotSpan sp = new SnapshotSpan(snapshot, intNext, currentSelection.Length);
                textview.Selection.Select(sp, false);
            }

            var initActive = textview.Selection.ActivePoint;

            SelectNextCommandFilter.m_trackList.RemoveAll(point => point.GetPosition(snapshot) == initActive.Position.Position);
            //var newAnchor = initAnchor.TranslateTo(textview.TextSnapshot, PointTrackingMode.Negative);
            var newActive = initActive.TranslateTo(textview.TextSnapshot, PointTrackingMode.Negative);
            //textview.Selection.Select(newAnchor, newActive);
            textview.Caret.MoveTo(newActive, PositionAffinity.Predecessor);
        }

    }
}
