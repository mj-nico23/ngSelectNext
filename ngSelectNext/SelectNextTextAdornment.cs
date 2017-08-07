using System;
using System.Windows.Controls;
using System.Windows.Media;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Editor;

namespace ngSelectNext
{

    internal sealed class SelectNextTextAdornment
    {
        private readonly IAdornmentLayer layer;

        private readonly IWpfTextView view;

        private readonly Brush brush;

        private readonly Pen pen;

        public SelectNextTextAdornment(IWpfTextView view)
        {
            if (view == null)
            {
                throw new ArgumentNullException("view");
            }

            this.layer = view.GetAdornmentLayer("SelectNextTextAdornment");

            this.view = view;
            this.view.LayoutChanged += this.OnLayoutChanged;

            // Create the pen and brush to color the box behind the a's
            this.brush = new SolidColorBrush(Color.FromArgb(0x20, 0x00, 0x00, 0xff));
            this.brush.Freeze();

            var penBrush = new SolidColorBrush(Colors.Blue);
            penBrush.Freeze();
            this.pen = new Pen(penBrush, 0.5);
            this.pen.Freeze();
        }

        internal void OnLayoutChanged(object sender, TextViewLayoutChangedEventArgs e)
        {
            this.CreateVisuals();
        }

        public void CreateVisuals()
        {
            IWpfTextViewLineCollection textViewLines = this.view.TextViewLines;

            if (MultiPointEditCommandFilter.m_trackList == null) return;

            string currentSelection = view.Selection.SelectedSpans[view.Selection.SelectedSpans.Count - 1].GetText();

            for (int i = 0; i < MultiPointEditCommandFilter.m_trackList.Count; i++)
            {
                int position = MultiPointEditCommandFilter.m_trackList[i].GetPosition(view.TextSnapshot);
                SnapshotSpan span = new SnapshotSpan(this.view.TextSnapshot, Span.FromBounds(position - currentSelection.Length, position));
                Geometry geometry = textViewLines.GetMarkerGeometry(span);
                if (geometry != null)
                {
                    var drawing = new GeometryDrawing(this.brush, this.pen, geometry);
                    drawing.Freeze();

                    var drawingImage = new DrawingImage(drawing);
                    drawingImage.Freeze();

                    var image = new Image
                    {
                        Source = drawingImage,
                    };

                    // Align the image with the top of the bounds of the text geometry
                    Canvas.SetLeft(image, geometry.Bounds.Left);
                    Canvas.SetTop(image, geometry.Bounds.Top);

                    this.layer.AddAdornment(AdornmentPositioningBehavior.TextRelative, span, null, image, null);
                }
            }
        }
    }
}
