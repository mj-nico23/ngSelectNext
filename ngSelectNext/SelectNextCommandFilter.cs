﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using EnvDTE;
using EnvDTE80;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;

namespace ngSelectNext
{
    internal class SelectNextCommandFilter : IOleCommandTarget
    {
        public static List<ITrackingPoint> m_trackList;

        private static bool addedCurrentCaret;
        private DTE2 m_dte;
        private static IWpfTextView m_textView;
        public static IAdornmentLayer m_adornmentLayer;
        private Dictionary<string, int> positionHash = new Dictionary<string, int>();

        public SelectNextCommandFilter(IWpfTextView tv)
        {
            m_textView = tv;
            m_adornmentLayer = tv.GetAdornmentLayer("SelectNextTextAdornment");
            m_dte = Microsoft.VisualStudio.Shell.ServiceProvider.GlobalProvider.GetService(typeof(DTE)) as DTE2;
        }

        public int QueryStatus(ref Guid pguidCmdGroup, uint cCmds, OLECMD[] prgCmds, IntPtr pCmdText)
        {
            if (pguidCmdGroup == typeof(VSConstants.VSStd2KCmdID).GUID)
            {
                for (int i = 0; i < cCmds; i++)
                {
                    switch (prgCmds[i].cmdID)
                    {
                        case ((uint)VSConstants.VSStd2KCmdID.TYPECHAR):
                        case ((uint)VSConstants.VSStd2KCmdID.BACKSPACE):
                        case ((uint)VSConstants.VSStd2KCmdID.TAB):
                        case ((uint)VSConstants.VSStd2KCmdID.LEFT):
                        case ((uint)VSConstants.VSStd2KCmdID.RIGHT):
                        case ((uint)VSConstants.VSStd2KCmdID.UP):
                        case ((uint)VSConstants.VSStd2KCmdID.DOWN):
                        case ((uint)VSConstants.VSStd2KCmdID.END):
                        case ((uint)VSConstants.VSStd2KCmdID.HOME):
                        case ((uint)VSConstants.VSStd2KCmdID.PAGEDN):
                        case ((uint)VSConstants.VSStd2KCmdID.PAGEUP):
                        case ((uint)VSConstants.VSStd2KCmdID.PASTE):
                        case ((uint)VSConstants.VSStd2KCmdID.PASTEASHTML):
                        case ((uint)VSConstants.VSStd2KCmdID.BOL):
                        case ((uint)VSConstants.VSStd2KCmdID.EOL):
                        case ((uint)VSConstants.VSStd2KCmdID.RETURN):
                        case ((uint)VSConstants.VSStd2KCmdID.BACKTAB):
                            prgCmds[i].cmdf = (uint)(OLECMDF.OLECMDF_ENABLED | OLECMDF.OLECMDF_SUPPORTED);
                            return VSConstants.S_OK;
                    }
                }
            }

            return NextTarget.QueryStatus(ref pguidCmdGroup, cCmds, prgCmds, pCmdText);
        }

        public int Exec(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            Debug.WriteLine(nCmdID);

            if (pguidCmdGroup == typeof(VSConstants.VSStd2KCmdID).GUID)
            {
                switch (nCmdID)
                {
                    case ((uint)VSConstants.VSStd2KCmdID.TYPECHAR):
                    case ((uint)VSConstants.VSStd2KCmdID.BACKSPACE):
                    case ((uint)VSConstants.VSStd2KCmdID.DELETEWORDRIGHT):
                    case ((uint)VSConstants.VSStd2KCmdID.DELETEWORDLEFT):
                    case ((uint)VSConstants.VSStd2KCmdID.TAB):
                    case ((uint)VSConstants.VSStd2KCmdID.LEFT):
                    case ((uint)VSConstants.VSStd2KCmdID.RIGHT):
                    case ((uint)VSConstants.VSStd2KCmdID.UP):
                    case ((uint)VSConstants.VSStd2KCmdID.DOWN):
                    case ((uint)VSConstants.VSStd2KCmdID.END):
                    case ((uint)VSConstants.VSStd2KCmdID.HOME):
                    case ((uint)VSConstants.VSStd2KCmdID.PAGEDN):
                    case ((uint)VSConstants.VSStd2KCmdID.PAGEUP):
                    case ((uint)VSConstants.VSStd2KCmdID.PASTE):
                    case ((uint)VSConstants.VSStd2KCmdID.PASTEASHTML):
                    case ((uint)VSConstants.VSStd2KCmdID.BOL):
                    case ((uint)VSConstants.VSStd2KCmdID.EOL):
                    case ((uint)VSConstants.VSStd2KCmdID.RETURN):
                    case ((uint)VSConstants.VSStd2KCmdID.BACKTAB):
                    case ((uint)VSConstants.VSStd2KCmdID.WORDPREV):
                    case ((uint)VSConstants.VSStd2KCmdID.WORDNEXT):


                        if (m_trackList?.Count > 0)
                            return SyncedOperation(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
                        break;
                }
            }
            else if (pguidCmdGroup == VSConstants.GUID_VSStandardCommandSet97)
            {
                switch (nCmdID)
                {

                    case ((uint)VSConstants.VSStd97CmdID.Delete):
                    case ((uint)VSConstants.VSStd97CmdID.Paste):
                        if (m_trackList?.Count > 0)
                            return SyncedOperation(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
                        break;
                }
            }


            switch (nCmdID)
            {

                // When ESC is used, cancel the Multi Edit mode
                case ((uint)VSConstants.VSStd2KCmdID.CANCEL):
                    ClearSyncPoints();
                    RedrawScreen();
                    break;
            }

            return NextTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
        }

        public static void ClearSyncPoints()
        {
            m_trackList?.Clear();
            addedCurrentCaret = false;
            m_adornmentLayer.RemoveAllAdornments();
        }

        private void RedrawScreen()
        {
            m_adornmentLayer.RemoveAllAdornments();
            positionHash.Clear();
            List<ITrackingPoint> newTrackList = new List<ITrackingPoint>();
            foreach (var trackPoint in m_trackList)
            {
                var curPosition = trackPoint.GetPosition(m_textView.TextSnapshot);
                IncrementCount(positionHash, curPosition.ToString());
                if (positionHash[curPosition.ToString()] > 1)
                    continue;
                DrawSingleSyncPoint(trackPoint);
                newTrackList.Add(trackPoint);
            }

            m_trackList = newTrackList;

        }

        private void IncrementCount(Dictionary<string, int> someDictionary, string id)
        {
            if (!someDictionary.ContainsKey(id))
                someDictionary[id] = 0;

            someDictionary[id]++;
        }

        private void DrawSingleSyncPoint(ITrackingPoint trackPoint)
        {
            if (trackPoint.GetPosition(m_textView.TextSnapshot) >= m_textView.TextSnapshot.Length)
                return;

            SnapshotSpan span = new SnapshotSpan(trackPoint.GetPoint(m_textView.TextSnapshot), 1);
            var brush = Brushes.DarkGray;
            var geom = m_textView.TextViewLines.GetLineMarkerGeometry(span);
            GeometryDrawing drawing = new GeometryDrawing(brush, null, geom);

            if (drawing.Bounds.IsEmpty)
                return;

            Rectangle rect = new Rectangle()
            {
                Fill = brush,
                Width = drawing.Bounds.Width / 6,
                Height = drawing.Bounds.Height - 4,
                Margin = new Thickness(0, 2, 0, 0),
            };

            Canvas.SetLeft(rect, geom.Bounds.Left);
            Canvas.SetTop(rect, geom.Bounds.Top);
            m_adornmentLayer.AddAdornment(AdornmentPositioningBehavior.TextRelative, span, "SelectNextLayer", rect, null);
        }

        private int SyncedOperation(ref Guid pguidCmdGroup, uint nCmdID, uint nCmdexecopt, IntPtr pvaIn, IntPtr pvaOut)
        {
            int result = 0;
            ITextCaret caret = m_textView.Caret;

            if (!addedCurrentCaret)
            {
                AddSyncPoint(caret.Position);
                addedCurrentCaret = true;
            }

            var tempTrackList = m_trackList;
            m_trackList = new List<ITrackingPoint>();

            m_dte.UndoContext.Open("Select Next edit");

            bool deleteSelection = !m_textView.Selection.IsEmpty;

            string currentSelection = m_textView.Selection.SelectedSpans[m_textView.Selection.SelectedSpans.Count - 1].GetText();

            for (int i = 0; i < tempTrackList.Count; i++)
            {
                var trackingMode = tempTrackList[i].TrackingMode;
                var snapPoint = tempTrackList[i].GetPoint(m_textView.TextSnapshot);
                caret.MoveTo(snapPoint);

                using (var edit = m_textView.TextSnapshot.TextBuffer.CreateEdit())
                {
                    bool applyEdit = false;
                    if (deleteSelection)
                    {
                        if (trackingMode == PointTrackingMode.Negative)
                            edit.Delete(caret.Position.BufferPosition.Position, currentSelection.Length);
                        else
                            edit.Delete(caret.Position.BufferPosition.Position - currentSelection.Length, currentSelection.Length);
                        applyEdit = true;
                    }

                    switch (nCmdID)
                    {
                        case (uint)VSConstants.VSStd2KCmdID.TYPECHAR:

                            var typedChar = (char)(ushort)Marshal.GetObjectForNativeVariant(pvaIn);
                            edit.Insert(tempTrackList[i].GetPosition(m_textView.TextSnapshot), typedChar.ToString());
                            applyEdit = true;
                            break;
                        case ((uint)VSConstants.VSStd97CmdID.Delete):
                            if (i < tempTrackList.Count - 1 && deleteSelection)
                            {
                                edit.Insert(tempTrackList[i].GetPosition(m_textView.TextSnapshot), "");
                            }
                            break;
                        case ((uint)VSConstants.VSStd97CmdID.Paste):
                            if (i < tempTrackList.Count - 1)
                            {
                                var pasteString = Clipboard.GetText();
                                edit.Insert(tempTrackList[i].GetPosition(m_textView.TextSnapshot), pasteString);
                                applyEdit = true;
                            }
                            break;
                    }

                    if (applyEdit)
                        edit.Apply();

                }


                switch (nCmdID)
                {
                    case ((uint)VSConstants.VSStd2KCmdID.TYPECHAR):
                        break;
                    case ((uint)VSConstants.VSStd97CmdID.Delete):
                        if (!deleteSelection)
                            result = NextTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
                        break;
                    case ((uint)VSConstants.VSStd97CmdID.Paste):
                        if (i == tempTrackList.Count - 1)
                        {
                            result = NextTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
                        }
                        break;
                    default:
                        result = NextTarget.Exec(ref pguidCmdGroup, nCmdID, nCmdexecopt, pvaIn, pvaOut);
                        break;
                }


                AddSyncPoint(m_textView.Caret.Position);
            }

            m_dte.UndoContext.Close();

            RedrawScreen();
            return result;
        }

        private void AddSyncPoint(CaretPosition caretPosition)
        {
            CaretPosition curPosition = caretPosition;
            // We don't support Virtual Spaces [yet?]

            var curTrackPoint = m_textView.TextSnapshot.CreateTrackingPoint(curPosition.BufferPosition.Position, PointTrackingMode.Positive);
            // Check if the bounds are valid

            if (curTrackPoint.GetPosition(m_textView.TextSnapshot) >= 0)
                m_trackList.Add(curTrackPoint);
            else
            {
                curTrackPoint = m_textView.TextSnapshot.CreateTrackingPoint(0, PointTrackingMode.Positive);
                m_trackList.Add(curTrackPoint);
            }

            if (curPosition.VirtualSpaces > 0)
            {
                m_textView.Caret.MoveTo(curTrackPoint.GetPoint(m_textView.TextSnapshot));
            }
        }

        internal bool Added { get; set; }
        internal IOleCommandTarget NextTarget { get; set; }
    }
}
