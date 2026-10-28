using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Imaging;

// Paints an EmfScene into a PdfGraphicsBuilder. The scene's frame (device pixels)
// maps onto the destination rectangle; every coordinate is converted to PDF space up
// front, so clipping and line state need nothing from the PDF transformation matrix.
internal static class EmfPdfWriter
{
    // Bitmaps are decoded and every coordinate validated before the first operator
    // is written, so a failure leaves the builder untouched. addImage registers a
    // page image and returns its resource name.
    public static void Draw(PdfGraphicsBuilder graphics, EmfScene scene, double x, double y, double width, double height,
        Func<PdfImageXObject, string> addImage, CancellationToken cancellationToken)
    {
        double scaleX = width / scene.FrameWidth;
        double scaleY = height / scene.FrameHeight;
        if (!double.IsFinite(scaleX) || !double.IsFinite(scaleY) || scaleX <= 0d || scaleY <= 0d)
        {
            throw new InvalidDataException("EMF destination rectangle is not usable.");
        }

        var bitmaps = new Dictionary<EmfBitmapOp, PdfImageXObject>();
        foreach (EmfOp op in scene.Ops)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (op is EmfBitmapOp bitmap)
            {
                using DecodedPixels pixels = scene.DecodeBitmap(bitmap, cancellationToken);
                bitmaps[bitmap] = PdfImageXObject.RgbPng(pixels.Width, pixels.Height, pixels.Rgb, pixels.Alpha);
                if (!IsFinite(bitmap.UnitToDevice))
                {
                    throw new InvalidDataException("EMF bitmap placement is not finite.");
                }
            }
            else if (op is EmfPathOp path && !path.Segments.TrueForAll(IsFinite))
            {
                throw new InvalidDataException("EMF path coordinates are not finite.");
            }

            for (EmfClip? clip = op.Clip; clip is not null; clip = clip.Parent)
            {
                if (!clip.Segments.TrueForAll(IsFinite))
                {
                    throw new InvalidDataException("EMF clip coordinates are not finite.");
                }
            }
        }

        double X(double deviceX) => x + (deviceX - scene.FrameLeft) * scaleX;
        double Y(double deviceY) => y + height - (deviceY - scene.FrameTop) * scaleY;
        double strokeScale = Math.Sqrt(scaleX * scaleY);

        graphics.SaveState();
        graphics.ClipRectangle(x, y, width, height);
        var appliedClips = new List<EmfClip>();
        bool dashed = false;
        foreach (EmfOp op in scene.Ops)
        {
            cancellationToken.ThrowIfCancellationRequested();
            SyncClips(op.Clip);
            if (op is EmfPathOp pathOp)
            {
                PaintPath(pathOp);
            }
            else if (op is EmfBitmapOp bitmapOp)
            {
                PaintBitmap(bitmapOp);
            }
        }

        SyncClips(null);
        graphics.RestoreState();

        void SyncClips(EmfClip? target)
        {
            var chain = new List<EmfClip>();
            for (EmfClip? clip = target; clip is not null; clip = clip.Parent)
            {
                chain.Add(clip);
            }

            chain.Reverse();

            int common = 0;
            while (common < appliedClips.Count && common < chain.Count && ReferenceEquals(appliedClips[common], chain[common]))
            {
                common++;
            }

            while (appliedClips.Count > common)
            {
                graphics.RestoreState();
                appliedClips.RemoveAt(appliedClips.Count - 1);
            }

            for (int i = common; i < chain.Count; i++)
            {
                graphics.SaveState();
                if (chain[i].Segments.Count == 0)
                {
                    graphics.ClipRectangle(0d, 0d, 0d, 0d);
                }
                else
                {
                    AppendPath(chain[i].Segments);
                    if (chain[i].EvenOdd)
                    {
                        graphics.ClipCurrentPathEvenOdd();
                    }
                    else
                    {
                        graphics.ClipCurrentPath();
                    }
                }

                appliedClips.Add(chain[i]);
            }
        }

        void AppendPath(List<EmfSegment> segments)
        {
            foreach (EmfSegment segment in segments)
            {
                switch (segment.Kind)
                {
                    case EmfSegmentKind.Move:
                        graphics.MoveTo(X(segment.X1), Y(segment.Y1));
                        break;
                    case EmfSegmentKind.Line:
                        graphics.LineTo(X(segment.X1), Y(segment.Y1));
                        break;
                    case EmfSegmentKind.Curve:
                        graphics.CurveTo(X(segment.X1), Y(segment.Y1), X(segment.X2), Y(segment.Y2), X(segment.X3), Y(segment.Y3));
                        break;
                    default:
                        graphics.ClosePath();
                        break;
                }
            }
        }

        void PaintPath(EmfPathOp op)
        {
            if (op.Stroke is { } stroke)
            {
                graphics.SetStrokeRgb(Red(stroke.Color), Green(stroke.Color), Blue(stroke.Color));
                graphics.SetLineWidth(Math.Max(stroke.Width * strokeScale, 0.01d));
                graphics.SetLineCap(stroke.Cap);
                graphics.SetLineJoin(stroke.Join);
                graphics.SetMiterLimit(stroke.MiterLimit);
                if (stroke.Dash is { Length: > 0 } dash)
                {
                    graphics.SetLineDash(dash.Select(length => Math.Max(length * strokeScale, 0.01d)).ToArray());
                    dashed = true;
                }
                else if (dashed)
                {
                    graphics.ClearLineDash();
                    dashed = false;
                }
            }

            if (op.FillColor is { } fill)
            {
                graphics.SetFillRgb(Red(fill), Green(fill), Blue(fill));
            }

            AppendPath(op.Segments);
            if (op.FillColor is not null && op.Stroke is not null)
            {
                if (op.EvenOdd) graphics.FillAndStrokeCurrentPathEvenOdd(); else graphics.FillAndStrokeCurrentPath();
            }
            else if (op.FillColor is not null)
            {
                if (op.EvenOdd) graphics.FillCurrentPathEvenOdd(); else graphics.FillCurrentPath();
            }
            else
            {
                graphics.StrokeCurrentPath();
            }
        }

        void PaintBitmap(EmfBitmapOp op)
        {
            EmfMatrix m = op.UnitToDevice;
            // The PDF image space has its origin at the bottom-left and v = 1 - t.
            double a = scaleX * m.A, b = -scaleY * m.B, c = -scaleX * m.C, d = scaleY * m.D;
            double e = scaleX * (m.C + m.E) + x - scene.FrameLeft * scaleX;
            double f = -scaleY * (m.D + m.F) + y + height + scene.FrameTop * scaleY;
            string name = addImage(bitmaps[op]);
            graphics.SaveState();
            graphics.Transform(a, b, c, d, e, f);
            graphics.DrawImage(name, 0d, 0d, 1d, 1d);
            graphics.RestoreState();
        }
    }

    private static bool IsFinite(EmfSegment s) =>
        double.IsFinite(s.X1) && double.IsFinite(s.Y1) && double.IsFinite(s.X2) && double.IsFinite(s.Y2) && double.IsFinite(s.X3) && double.IsFinite(s.Y3);

    private static bool IsFinite(EmfMatrix m) =>
        double.IsFinite(m.A) && double.IsFinite(m.B) && double.IsFinite(m.C) && double.IsFinite(m.D) && double.IsFinite(m.E) && double.IsFinite(m.F);

    private static byte Red(uint color) => (byte)(color & 0xFF);

    private static byte Green(uint color) => (byte)((color >> 8) & 0xFF);

    private static byte Blue(uint color) => (byte)((color >> 16) & 0xFF);
}
