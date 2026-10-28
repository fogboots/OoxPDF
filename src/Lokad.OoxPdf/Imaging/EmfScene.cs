namespace Lokad.OoxPdf.Imaging;

// Enhanced Metafile playback into a device-space display list (MS-EMF). Records are
// played against a GDI-style device context (mapping mode, world transform, clip,
// pen/brush objects, path brackets) and reduced to filled/stroked paths and DIB
// bitmaps whose coordinates are device pixels of the reference device. Anything
// that cannot be reproduced faithfully (arcs, raster operations, regions other
// than rectangles, text, EMF+-only files) throws NotSupportedException so the
// caller keeps its IMAGE_UNSUPPORTED_FORMAT diagnostic. EMF+ dual files are
// rendered from their EMF records and the EMF+ comments are ignored.

internal readonly record struct EmfMatrix(double A, double B, double C, double D, double E, double F)
{
    public static EmfMatrix Identity => new(1d, 0d, 0d, 1d, 0d, 0d);

    public (double X, double Y) Apply(double x, double y) => (A * x + C * y + E, B * x + D * y + F);

    // Applies this transform first, then <paramref name="next"/>.
    public EmfMatrix Then(EmfMatrix next) => new(
        A * next.A + B * next.C, A * next.B + B * next.D,
        C * next.A + D * next.C, C * next.B + D * next.D,
        E * next.A + F * next.C + next.E, E * next.B + F * next.D + next.F);

    public double ScaleFactor => Math.Sqrt(Math.Abs(A * D - B * C));
}

internal enum EmfSegmentKind { Move, Line, Curve, Close }

internal readonly record struct EmfSegment(EmfSegmentKind Kind, double X1, double Y1, double X2, double Y2, double X3, double Y3);

internal sealed class EmfClip(EmfClip? parent, List<EmfSegment> segments, bool evenOdd)
{
    // The effective clip is the intersection of this region with every parent.
    public EmfClip? Parent { get; } = parent;

    // Each level costs one nested PDF graphics state.
    public int Depth { get; } = (parent?.Depth ?? 0) + 1;

    public List<EmfSegment> Segments { get; } = segments;

    public bool EvenOdd { get; } = evenOdd;
}

internal sealed record EmfStroke(uint Color, double Width, double[]? Dash, int Cap, int Join, double MiterLimit);

internal abstract class EmfOp
{
    public EmfClip? Clip { get; init; }
}

internal sealed class EmfPathOp : EmfOp
{
    public required List<EmfSegment> Segments { get; init; }

    public bool EvenOdd { get; init; }

    public uint? FillColor { get; init; }

    public EmfStroke? Stroke { get; init; }
}

internal sealed class EmfBitmapOp : EmfOp
{
    // Maps the unit square of the cropped bitmap (t = 0 at its top edge) to device pixels.
    public required EmfMatrix UnitToDevice { get; init; }

    public required int InfoOffset { get; init; }

    public required int InfoLength { get; init; }

    public required int BitsOffset { get; init; }

    public required int BitsLength { get; init; }

    public required int SrcX { get; init; }

    public required int SrcY { get; init; }

    public required int SrcWidth { get; init; }

    public required int SrcHeight { get; init; }
}

internal sealed class EmfScene
{
    private const uint SignatureValue = 0x464D4520u;

    private EmfScene(byte[] data, double frameLeft, double frameTop, double frameWidth, double frameHeight, List<EmfOp> ops)
    {
        Data = data;
        FrameLeft = frameLeft;
        FrameTop = frameTop;
        FrameWidth = frameWidth;
        FrameHeight = frameHeight;
        Ops = ops;
    }

    public byte[] Data { get; }

    // The picture frame (rclFrame) in device pixels; it maps onto the destination rectangle.
    public double FrameLeft { get; }

    public double FrameTop { get; }

    public double FrameWidth { get; }

    public double FrameHeight { get; }

    public IReadOnlyList<EmfOp> Ops { get; }

    // True when the metafile is one unclipped bitmap that fills the frame: it is then
    // an ordinary raster picture and shares the pixel pipeline (crop, recolor, cache).
    public bool IsBitmapOnly => TryGetFrameBitmap(out _);

    public static bool IsEmf(byte[] bytes)
    {
        return bytes.Length >= 88 && U32(bytes, 0) == 1u && U32(bytes, 40) == SignatureValue;
    }

    public static EmfScene Read(byte[] data, CancellationToken cancellationToken = default)
    {
        if (!IsEmf(data))
        {
            throw new InvalidDataException("Data is not an EMF image.");
        }

        try
        {
            return new EmfPlayer(data, cancellationToken).Play();
        }
        catch (Exception ex) when (ex is ArgumentException or IndexOutOfRangeException or OverflowException)
        {
            throw new InvalidDataException("EMF record data is truncated or inconsistent.", ex);
        }
    }

    public DecodedPixels DecodeSingleBitmap(CancellationToken cancellationToken = default)
    {
        if (!TryGetFrameBitmap(out EmfBitmapOp? bitmap))
        {
            throw new NotSupportedException("EMF content is not a single bitmap and cannot be decoded as pixels.");
        }

        DecodedPixels pixels = DecodeBitmap(bitmap, cancellationToken);
        // Device-space orientation: a negative scale flips the cropped bitmap.
        bool flipX = bitmap.UnitToDevice.A < 0d;
        bool flipY = bitmap.UnitToDevice.D < 0d;
        if (!flipX && !flipY)
        {
            return pixels;
        }

        using (pixels)
        {
            int width = pixels.Width, height = pixels.Height;
            byte[] rgb = new byte[pixels.Rgb.Length];
            byte[]? alpha = pixels.Alpha is null ? null : new byte[pixels.Alpha.Length];
            for (int y = 0; y < height; y++)
            {
                int sourceRow = (flipY ? height - 1 - y : y) * width;
                for (int x = 0; x < width; x++)
                {
                    int source = sourceRow + (flipX ? width - 1 - x : x);
                    int target = y * width + x;
                    rgb[target * 3] = pixels.Rgb[source * 3];
                    rgb[target * 3 + 1] = pixels.Rgb[source * 3 + 1];
                    rgb[target * 3 + 2] = pixels.Rgb[source * 3 + 2];
                    if (alpha is not null)
                    {
                        alpha[target] = pixels.Alpha![source];
                    }
                }
            }

            return new DecodedPixels(width, height, rgb, alpha);
        }
    }

    // Decodes the DIB and crops it to the record's source rectangle.
    public DecodedPixels DecodeBitmap(EmfBitmapOp bitmap, CancellationToken cancellationToken = default)
    {
        using DecodedPixels whole = DibImage.ReadOwned(Data, bitmap.InfoOffset, bitmap.InfoLength, bitmap.BitsOffset, bitmap.BitsLength, 0, cancellationToken);
        int x0 = bitmap.SrcX, y0 = bitmap.SrcY, width = bitmap.SrcWidth, height = bitmap.SrcHeight;
        if (x0 < 0 || y0 < 0 || width <= 0 || height <= 0 || x0 + width > whole.Width || y0 + height > whole.Height)
        {
            // Embedded PNG/JPEG payloads may differ in size from the declared header.
            throw new InvalidDataException("EMF bitmap source rectangle lies outside the bitmap.");
        }

        if (x0 == 0 && y0 == 0 && width == whole.Width && height == whole.Height)
        {
            return new DecodedPixels(width, height, whole.Rgb, whole.Alpha);
        }

        byte[] rgb = new byte[width * height * 3];
        byte[]? alpha = whole.Alpha is null ? null : new byte[width * height];
        for (int y = 0; y < height; y++)
        {
            Array.Copy(whole.Rgb, ((y0 + y) * whole.Width + x0) * 3, rgb, y * width * 3, width * 3);
            if (alpha is not null)
            {
                Array.Copy(whole.Alpha!, (y0 + y) * whole.Width + x0, alpha, y * width, width);
            }
        }

        return new DecodedPixels(width, height, rgb, alpha);
    }

    private bool TryGetFrameBitmap(out EmfBitmapOp bitmap)
    {
        bitmap = null!;
        if (Ops.Count != 1 || Ops[0] is not EmfBitmapOp candidate || candidate.Clip is not null)
        {
            return false;
        }

        EmfMatrix m = candidate.UnitToDevice;
        if (Math.Abs(m.B) > 1e-6 || Math.Abs(m.C) > 1e-6 || m.A == 0d || m.D == 0d)
        {
            return false;
        }

        double left = Math.Min(m.E, m.E + m.A), right = Math.Max(m.E, m.E + m.A);
        double top = Math.Min(m.F, m.F + m.D), bottom = Math.Max(m.F, m.F + m.D);
        double toleranceX = Math.Max(2d, FrameWidth * 0.015d);
        double toleranceY = Math.Max(2d, FrameHeight * 0.015d);
        if (Math.Abs(left - FrameLeft) > toleranceX || Math.Abs(right - (FrameLeft + FrameWidth)) > toleranceX ||
            Math.Abs(top - FrameTop) > toleranceY || Math.Abs(bottom - (FrameTop + FrameHeight)) > toleranceY)
        {
            return false;
        }

        bitmap = candidate;
        return true;
    }

    private static uint U32(byte[] bytes, int offset)
    {
        return BitConverter.ToUInt32(bytes, offset);
    }

    private sealed record PenObject(bool IsNull, uint Color, double Width, int Style, int Cap, int Join, bool Geometric, double[]? UserDash);

    private sealed record BrushObject(bool IsNull, uint Color, string? Unsupported);

    private sealed class DeviceContext
    {
        public int MapMode = 1;
        public double WindowOrgX, WindowOrgY, ViewportOrgX, ViewportOrgY;
        public double WindowExtX = 1d, WindowExtY = 1d, ViewportExtX = 1d, ViewportExtY = 1d;
        public EmfMatrix World = EmfMatrix.Identity;
        public object Pen = BlackPen;
        public object Brush = WhiteBrush;
        public int PolyFillMode = 1;
        public int Rop2 = 13;
        public double MiterLimit = 10d;
        public EmfClip? Clip;

        public DeviceContext Clone() => (DeviceContext)MemberwiseClone();
    }

    private static readonly PenObject BlackPen = new(false, 0u, 0d, 0, 1, 1, false, null);
    private static readonly BrushObject WhiteBrush = new(false, 0x00FFFFFFu, null);

    private sealed class EmfPlayer(byte[] data, CancellationToken cancellationToken)
    {
        private readonly List<EmfOp> ops = [];
        private readonly Dictionary<int, object> objects = [];
        private readonly List<DeviceContext> savedContexts = [];
        private DeviceContext dc = new();
        private double pixelsPerMmX = 1d, pixelsPerMmY = 1d;
        private double currentX, currentY;
        // The open path bracket (BEGINPATH..ENDPATH) and the finished path awaiting a
        // fill/stroke/clip record.
        private List<EmfSegment>? path;
        private List<EmfSegment>? finishedPath;
        private bool figureOpen;
        private double figureStartX, figureStartY;
        private bool emfPlusOnly;

        public EmfScene Play()
        {
            int deviceWidth = I32(72), deviceHeight = I32(76), millimetersWidth = I32(80), millimetersHeight = I32(84);
            if (deviceWidth > 0 && deviceHeight > 0 && millimetersWidth > 0 && millimetersHeight > 0)
            {
                pixelsPerMmX = (double)deviceWidth / millimetersWidth;
                pixelsPerMmY = (double)deviceHeight / millimetersHeight;
            }

            double frameLeft = I32(24) / 100d * pixelsPerMmX, frameTop = I32(28) / 100d * pixelsPerMmY;
            double frameWidth = (I32(32) - I32(24)) / 100d * pixelsPerMmX, frameHeight = (I32(36) - I32(28)) / 100d * pixelsPerMmY;
            if (!(frameWidth > 0d) || !(frameHeight > 0d) || deviceWidth <= 0 || millimetersWidth <= 0)
            {
                // No usable frame: fall back to the inclusive device bounds.
                frameLeft = I32(8);
                frameTop = I32(12);
                frameWidth = I32(16) - I32(8) + 1d;
                frameHeight = I32(20) - I32(12) + 1d;
            }

            if (!double.IsFinite(frameWidth) || !double.IsFinite(frameHeight) || frameWidth <= 0d || frameHeight <= 0d ||
                !double.IsFinite(frameLeft) || !double.IsFinite(frameTop))
            {
                throw new InvalidDataException("EMF frame is not usable.");
            }

            int position = U32Int(4);
            int count = 0;
            while (position + 8 <= data.Length)
            {
                if ((++count & 1023) == 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                }

                int type = I32(position);
                int size = I32(position + 4);
                if (size < 8 || (size & 3) != 0 || position + (long)size > data.Length)
                {
                    throw new InvalidDataException("EMF record extends beyond the end of the data.");
                }

                if (type == 14)
                {
                    break;
                }

                PlayRecord(type, position, size);
                position += size;
            }

            if (emfPlusOnly)
            {
                throw new NotSupportedException("EMF+-only metafile has no EMF drawing records.");
            }

            return new EmfScene(data, frameLeft, frameTop, frameWidth, frameHeight, ops);
        }

        private void PlayRecord(int type, int p, int size)
        {
            switch (type)
            {
                case 70: // EMR_COMMENT
                    ReadComment(p, size);
                    break;
                case 9: dc.WindowExtX = I32(p + 8); dc.WindowExtY = I32(p + 12); break;
                case 10: dc.WindowOrgX = I32(p + 8); dc.WindowOrgY = I32(p + 12); break;
                case 11: dc.ViewportExtX = I32(p + 8); dc.ViewportExtY = I32(p + 12); break;
                case 12: dc.ViewportOrgX = I32(p + 8); dc.ViewportOrgY = I32(p + 12); break;
                case 31: // EMR_SCALEVIEWPORTEXTEX
                    dc.ViewportExtX = Scale(dc.ViewportExtX, I32(p + 8), I32(p + 12));
                    dc.ViewportExtY = Scale(dc.ViewportExtY, I32(p + 16), I32(p + 20));
                    break;
                case 32: // EMR_SCALEWINDOWEXTEX
                    dc.WindowExtX = Scale(dc.WindowExtX, I32(p + 8), I32(p + 12));
                    dc.WindowExtY = Scale(dc.WindowExtY, I32(p + 16), I32(p + 20));
                    break;
                case 17: dc.MapMode = I32(p + 8); break;
                case 19: dc.PolyFillMode = I32(p + 8); break;
                case 20: dc.Rop2 = I32(p + 8); break;
                case 58: dc.MiterLimit = Math.Max(1d, U32Int(p + 8)); break;
                case 35: dc.World = ReadTransform(p + 8); break;
                case 36: ModifyWorldTransform(p); break;
                case 33: savedContexts.Add(dc.Clone()); break;
                case 34: RestoreContext(I32(p + 8)); break;
                case 37: SelectObject(I32(p + 8)); break;
                case 38: CreatePen(p); break;
                case 95: CreateExtPen(p, size); break;
                case 39: CreateBrush(p); break;
                case 40: objects.Remove(I32(p + 8)); break;
                case 27: // EMR_MOVETOEX
                    currentX = I32(p + 8);
                    currentY = I32(p + 12);
                    figureOpen = false;
                    if (path is not null)
                    {
                        StartFigure(path);
                    }
                    break;
                case 54: LineTo(I32(p + 8), I32(p + 12)); break;
                case 59: path = []; finishedPath = null; figureOpen = false; break; // EMR_BEGINPATH
                case 60: finishedPath = path; path = null; figureOpen = false; break; // EMR_ENDPATH
                case 61: CloseFigure(); break;
                case 62: PaintPath(fill: true, stroke: false); break;
                case 63: PaintPath(fill: true, stroke: true); break;
                case 64: PaintPath(fill: false, stroke: true); break;
                case 67: SelectClipPath(I32(p + 8)); break;
                case 30: // EMR_INTERSECTCLIPRECT
                    PushClip(MapToDevice(RectanglePath(I32(p + 8), I32(p + 12), I32(p + 16), I32(p + 20))), evenOdd: false);
                    break;
                case 75: ExtSelectClipRegion(p); break;
                case 3: PolyPoints(p, size, wide: true, EmfPolyKind.Polygon); break;
                case 4: PolyPoints(p, size, wide: true, EmfPolyKind.Polyline); break;
                case 2: PolyPoints(p, size, wide: true, EmfPolyKind.Bezier); break;
                case 6: PolyPoints(p, size, wide: true, EmfPolyKind.PolylineTo); break;
                case 5: PolyPoints(p, size, wide: true, EmfPolyKind.BezierTo); break;
                case 86: PolyPoints(p, size, wide: false, EmfPolyKind.Polygon); break;
                case 87: PolyPoints(p, size, wide: false, EmfPolyKind.Polyline); break;
                case 85: PolyPoints(p, size, wide: false, EmfPolyKind.Bezier); break;
                case 89: PolyPoints(p, size, wide: false, EmfPolyKind.PolylineTo); break;
                case 88: PolyPoints(p, size, wide: false, EmfPolyKind.BezierTo); break;
                case 7: PolyPoly(p, size, wide: true, closed: false); break;
                case 8: PolyPoly(p, size, wide: true, closed: true); break;
                case 90: PolyPoly(p, size, wide: false, closed: false); break;
                case 91: PolyPoly(p, size, wide: false, closed: true); break;
                case 43: Shape(RectanglePath(I32(p + 8), I32(p + 12), I32(p + 16), I32(p + 20))); break;
                case 42: Shape(EllipsePath(I32(p + 8), I32(p + 12), I32(p + 16), I32(p + 20))); break;
                case 44: Shape(RoundRectanglePath(I32(p + 8), I32(p + 12), I32(p + 16), I32(p + 20), I32(p + 24), I32(p + 28))); break;
                case 81: StretchDibits(p); break;
                case 76: BitBlt(p, stretch: false); break;
                case 77: BitBlt(p, stretch: true); break;
                case 84: throw new NotSupportedException("EMF text records (EMR_EXTTEXTOUTW) are not supported.");
                case 93: case 94:
                    objects[I32(p + 8)] = new BrushObject(false, 0u, "pattern brushes");
                    break;
                case 1: case 13: case 16: case 18: case 21: case 22: case 23: case 24: case 25: case 28: case 48: case 49: case 50: case 51: case 52:
                case 57: case 82: case 98: case 99: case 100: case 101: case 109: case 111: case 112: case 113: case 115: case 119: case 120:
                    // Header, brush origin, background/text state, fonts (only text records use them),
                    // palette plumbing, ICM and layout: none changes the geometry or colors this player paints.
                    break;
                default:
                    throw new NotSupportedException($"EMF record type {type} is not supported.");
            }
        }

        private void ReadComment(int p, int size)
        {
            // EMR_COMMENT: byte count, then data. "EMF+" marks an EMF+ record stream;
            // its header record says whether the plain EMF records duplicate the drawing.
            if (size >= 28 && I32(p + 8) >= 16 && I32(p + 12) == 0x2B464D45)
            {
                int emfPlusType = data[p + 16] | (data[p + 17] << 8);
                int flags = data[p + 18] | (data[p + 19] << 8);
                if (emfPlusType == 0x4001 && (flags & 1) == 0)
                {
                    emfPlusOnly = true;
                }
            }
        }

        private static double Scale(double extent, int numerator, int denominator)
        {
            return denominator == 0 ? extent : extent * numerator / denominator;
        }

        private void RestoreContext(int relative)
        {
            int steps = relative < 0 ? -relative : savedContexts.Count - relative + 1;
            if (steps < 1 || steps > savedContexts.Count)
            {
                return;
            }

            dc = savedContexts[^steps];
            savedContexts.RemoveRange(savedContexts.Count - steps, steps);
        }

        private EmfMatrix ReadTransform(int p)
        {
            return new EmfMatrix(F32(p), F32(p + 4), F32(p + 8), F32(p + 12), F32(p + 16), F32(p + 20));
        }

        private void ModifyWorldTransform(int p)
        {
            EmfMatrix transform = ReadTransform(p + 8);
            dc.World = I32(p + 32) switch
            {
                1 => EmfMatrix.Identity,
                2 => transform.Then(dc.World),
                3 => dc.World.Then(transform),
                4 => transform,
                _ => throw new NotSupportedException("EMF world transform mode is not supported.")
            };
        }

        // Logical coordinates to reference-device pixels.
        private EmfMatrix LogicalToDevice()
        {
            double scaleX, scaleY;
            switch (dc.MapMode)
            {
                case 1:
                    scaleX = 1d;
                    scaleY = 1d;
                    break;
                case 7 or 8:
                    scaleX = dc.WindowExtX == 0d ? 1d : dc.ViewportExtX / dc.WindowExtX;
                    scaleY = dc.WindowExtY == 0d ? 1d : dc.ViewportExtY / dc.WindowExtY;
                    break;
                case >= 2 and <= 6:
                    double millimeters = dc.MapMode switch { 2 => 0.1d, 3 => 0.01d, 4 => 0.254d, 5 => 0.0254d, _ => 25.4d / 1440d };
                    scaleX = millimeters * pixelsPerMmX;
                    scaleY = -millimeters * pixelsPerMmY;
                    break;
                default:
                    throw new NotSupportedException("EMF mapping mode is not supported.");
            }

            var page = new EmfMatrix(scaleX, 0d, 0d, scaleY, dc.ViewportOrgX - scaleX * dc.WindowOrgX, dc.ViewportOrgY - scaleY * dc.WindowOrgY);
            return dc.World.Then(page);
        }

        private void SelectObject(int handle)
        {
            if ((handle & unchecked((int)0x80000000)) != 0)
            {
                switch (handle & 0x7FFFFFFF)
                {
                    case 0: dc.Brush = WhiteBrush; break;
                    case 1: dc.Brush = new BrushObject(false, 0x00C0C0C0u, null); break;
                    case 2: dc.Brush = new BrushObject(false, 0x00808080u, null); break;
                    case 3: dc.Brush = new BrushObject(false, 0x00404040u, null); break;
                    case 4: dc.Brush = new BrushObject(false, 0u, null); break;
                    case 5: dc.Brush = new BrushObject(true, 0u, null); break;
                    case 6: dc.Pen = new PenObject(false, 0x00FFFFFFu, 0d, 0, 1, 1, false, null); break;
                    case 7: dc.Pen = BlackPen; break;
                    case 8: dc.Pen = new PenObject(true, 0u, 0d, 5, 1, 1, false, null); break;
                    case 18: dc.Brush = WhiteBrush; break; // DC_BRUSH keeps its white default
                    case 19: dc.Pen = BlackPen; break; // DC_PEN keeps its black default
                }

                return;
            }

            if (objects.TryGetValue(handle, out object? selected) && selected is PenObject or BrushObject)
            {
                if (selected is PenObject)
                {
                    dc.Pen = selected;
                }
                else
                {
                    dc.Brush = selected;
                }
            }
        }

        private void CreatePen(int p)
        {
            int style = I32(p + 12) & 0xFF;
            objects[I32(p + 8)] = new PenObject(style == 5, (uint)I32(p + 24), I32(p + 16), style, 1, 1, false, null);
        }

        private void CreateExtPen(int p, int size)
        {
            int style = I32(p + 28);
            int type = style & 0xF;
            int cap = (style & 0xF00) switch { 0x100 => 2, 0x200 => 0, _ => 1 };
            int join = (style & 0xF000) switch { 0x1000 => 2, 0x2000 => 0, _ => 1 };
            bool geometric = (style & 0x10000) != 0;
            int brushStyle = I32(p + 36);
            if (brushStyle is not (0 or 1))
            {
                // Patterned and hatched pens have no solid color to stroke with.
                objects[I32(p + 8)] = new PenObject(false, uint.MaxValue, 0d, type, cap, join, geometric, null);
                return;
            }

            double[]? dashes = null;
            int dashCount = I32(p + 48);
            if (type == 7 && dashCount > 0 && dashCount <= 16 && p + 52 + dashCount * 4 <= p + size)
            {
                dashes = new double[dashCount];
                for (int i = 0; i < dashCount; i++)
                {
                    dashes[i] = U32Int(p + 52 + i * 4);
                }
            }

            objects[I32(p + 8)] = new PenObject(brushStyle == 1 || type == 5, (uint)I32(p + 40), U32Int(p + 32), type, cap, join, geometric, dashes);
        }

        private void CreateBrush(int p)
        {
            int style = I32(p + 12);
            string? unsupported = style switch
            {
                0 or 1 => null,
                2 => "hatched brushes",
                _ => "pattern brushes"
            };
            objects[I32(p + 8)] = new BrushObject(style == 1, (uint)I32(p + 16), unsupported);
        }

        private EmfStroke? CurrentStroke()
        {
            var pen = (PenObject)dc.Pen;
            if (pen.IsNull || pen.Style == 5)
            {
                return null;
            }

            if (pen.Color == uint.MaxValue || (pen.Color & 0x01000000u) != 0u)
            {
                throw new NotSupportedException("EMF pen uses a pattern or palette color.");
            }

            RequireCopyRop();
            double scale = LogicalToDevice().ScaleFactor;
            // Cosmetic and zero-width pens are one device pixel wide.
            double width = pen.Geometric || pen.Width > 1d ? Math.Max(pen.Width * scale, 1d) : 1d;
            double[]? dash = pen.Style switch
            {
                1 => [3d * width, width],
                2 => [width, width],
                3 => [3d * width, width, width, width],
                4 => [3d * width, width, width, width, width, width],
                7 when pen.UserDash is { Length: > 0 } user => user.Select(value => Math.Max(value * scale, 0.1d)).ToArray(),
                _ => null
            };
            return new EmfStroke(pen.Color, width, dash, pen.Cap, pen.Join, dc.MiterLimit);
        }

        private uint? CurrentFill()
        {
            var brush = (BrushObject)dc.Brush;
            if (brush.IsNull)
            {
                return null;
            }

            if (brush.Unsupported is not null)
            {
                throw new NotSupportedException("EMF " + brush.Unsupported + " are not supported.");
            }

            if ((brush.Color & 0x01000000u) != 0u)
            {
                throw new NotSupportedException("EMF brush uses a palette color.");
            }

            RequireCopyRop();
            return brush.Color;
        }

        private void RequireCopyRop()
        {
            if (dc.Rop2 != 13)
            {
                throw new NotSupportedException("EMF binary raster operations other than copy are not supported.");
            }
        }

        private (double X, double Y) ToDevice(double x, double y)
        {
            return LogicalToDevice().Apply(x, y);
        }

        // Begins a figure at the current position (logical coordinates).
        private void StartFigure(List<EmfSegment> target)
        {
            (double x, double y) = ToDevice(currentX, currentY);
            target.Add(new EmfSegment(EmfSegmentKind.Move, x, y, 0d, 0d, 0d, 0d));
            figureOpen = true;
            figureStartX = currentX;
            figureStartY = currentY;
        }

        private void AddLine(double logicalX, double logicalY, List<EmfSegment> target)
        {
            (double x, double y) = ToDevice(logicalX, logicalY);
            target.Add(new EmfSegment(EmfSegmentKind.Line, x, y, 0d, 0d, 0d, 0d));
        }

        private void LineTo(int x, int y)
        {
            List<EmfSegment> target = path ?? [];
            if (!figureOpen || path is null)
            {
                StartFigure(target);
            }

            AddLine(x, y, target);
            currentX = x;
            currentY = y;
            if (path is null)
            {
                figureOpen = false;
                Emit(target, fill: false, stroke: true);
            }
        }

        private void CloseFigure()
        {
            if (path is not null && figureOpen)
            {
                path.Add(new EmfSegment(EmfSegmentKind.Close, 0d, 0d, 0d, 0d, 0d, 0d));
                figureOpen = false;
                currentX = figureStartX;
                currentY = figureStartY;
            }
        }

        private enum EmfPolyKind { Polygon, Polyline, Bezier, PolylineTo, BezierTo }

        private (int X, int Y)[] ReadPoints(int offset, int count, bool wide)
        {
            var points = new (int X, int Y)[count];
            for (int i = 0; i < count; i++)
            {
                points[i] = wide
                    ? (I32(offset + i * 8), I32(offset + i * 8 + 4))
                    : ((short)(data[offset + i * 4] | (data[offset + i * 4 + 1] << 8)), (short)(data[offset + i * 4 + 2] | (data[offset + i * 4 + 3] << 8)));
            }

            return points;
        }

        private void PolyPoints(int p, int size, bool wide, EmfPolyKind kind)
        {
            int count = I32(p + 24);
            if (count < 0 || 28L + (long)count * (wide ? 8 : 4) > size)
            {
                throw new InvalidDataException("EMF polygon record is truncated.");
            }

            if (count == 0)
            {
                return;
            }

            (int X, int Y)[] points = ReadPoints(p + 28, count, wide);
            List<EmfSegment> target = path ?? [];
            bool toKind = kind is EmfPolyKind.PolylineTo or EmfPolyKind.BezierTo;
            int index = 0;
            if (toKind)
            {
                if (!figureOpen || path is null)
                {
                    StartFigure(target);
                }
            }
            else
            {
                (double startX, double startY) = ToDevice(points[0].X, points[0].Y);
                target.Add(new EmfSegment(EmfSegmentKind.Move, startX, startY, 0d, 0d, 0d, 0d));
                index = 1;
            }

            if (kind is EmfPolyKind.Bezier or EmfPolyKind.BezierTo)
            {
                for (; index + 2 < count; index += 3)
                {
                    AddCurve(target, points[index], points[index + 1], points[index + 2]);
                }
            }
            else
            {
                for (; index < count; index++)
                {
                    AddLine(points[index].X, points[index].Y, target);
                }
            }

            if (kind == EmfPolyKind.Polygon)
            {
                target.Add(new EmfSegment(EmfSegmentKind.Close, 0d, 0d, 0d, 0d, 0d, 0d));
            }

            if (toKind)
            {
                currentX = points[^1].X;
                currentY = points[^1].Y;
            }

            if (path is not null)
            {
                figureOpen = kind != EmfPolyKind.Polygon && (toKind || kind is EmfPolyKind.Polyline or EmfPolyKind.Bezier);
                return;
            }

            figureOpen = false;
            Emit(target, fill: kind == EmfPolyKind.Polygon, stroke: true);
        }

        private void AddCurve(List<EmfSegment> target, (int X, int Y) c1, (int X, int Y) c2, (int X, int Y) end)
        {
            (double x1, double y1) = ToDevice(c1.X, c1.Y);
            (double x2, double y2) = ToDevice(c2.X, c2.Y);
            (double x3, double y3) = ToDevice(end.X, end.Y);
            target.Add(new EmfSegment(EmfSegmentKind.Curve, x1, y1, x2, y2, x3, y3));
        }

        private void PolyPoly(int p, int size, bool wide, bool closed)
        {
            int polygons = I32(p + 24);
            int total = I32(p + 28);
            int countsOffset = p + 32;
            int pointsOffset = countsOffset + polygons * 4;
            if (polygons < 0 || total < 0 || pointsOffset + (long)total * (wide ? 8 : 4) > p + size)
            {
                throw new InvalidDataException("EMF poly-polygon record is truncated.");
            }

            (int X, int Y)[] points = ReadPoints(pointsOffset, total, wide);
            List<EmfSegment> target = path ?? [];
            int start = 0;
            for (int polygon = 0; polygon < polygons; polygon++)
            {
                int count = I32(countsOffset + polygon * 4);
                if (count < 0 || start + count > total)
                {
                    throw new InvalidDataException("EMF poly-polygon counts are inconsistent.");
                }

                for (int i = 0; i < count; i++)
                {
                    (double x, double y) = ToDevice(points[start + i].X, points[start + i].Y);
                    target.Add(new EmfSegment(i == 0 ? EmfSegmentKind.Move : EmfSegmentKind.Line, x, y, 0d, 0d, 0d, 0d));
                }

                if (closed && count > 0)
                {
                    target.Add(new EmfSegment(EmfSegmentKind.Close, 0d, 0d, 0d, 0d, 0d, 0d));
                }

                start += count;
            }

            if (path is null)
            {
                Emit(target, fill: closed, stroke: true);
            }
            else
            {
                figureOpen = false;
            }
        }

        private void Shape(List<EmfSegment> logicalShape)
        {
            // Shapes are authored in logical units and mapped here, so rotated and
            // sheared world transforms keep the curve control points exact.
            List<EmfSegment> mapped = MapToDevice(logicalShape);
            if (path is null)
            {
                Emit(mapped, fill: true, stroke: true);
                return;
            }

            path.AddRange(mapped);
            figureOpen = false;
        }

        private static List<EmfSegment> RectanglePath(double left, double top, double right, double bottom)
        {
            return
            [
                new(EmfSegmentKind.Move, left, top, 0d, 0d, 0d, 0d),
                new(EmfSegmentKind.Line, right, top, 0d, 0d, 0d, 0d),
                new(EmfSegmentKind.Line, right, bottom, 0d, 0d, 0d, 0d),
                new(EmfSegmentKind.Line, left, bottom, 0d, 0d, 0d, 0d),
                new(EmfSegmentKind.Close, 0d, 0d, 0d, 0d, 0d, 0d)
            ];
        }

        private const double Kappa = 0.5522847498307936d;

        private static List<EmfSegment> EllipsePath(double left, double top, double right, double bottom)
        {
            return RoundedPath(left, top, right, bottom, (right - left) / 2d, (bottom - top) / 2d);
        }

        private static List<EmfSegment> RoundRectanglePath(double left, double top, double right, double bottom, double cornerWidth, double cornerHeight)
        {
            double rx = Math.Min(Math.Abs(cornerWidth) / 2d, Math.Abs(right - left) / 2d);
            double ry = Math.Min(Math.Abs(cornerHeight) / 2d, Math.Abs(bottom - top) / 2d);
            return RoundedPath(left, top, right, bottom, rx, ry);
        }

        // A rectangle whose corners are quarter ellipses of radii rx, ry (a full ellipse
        // when the radii are half the extents).
        private static List<EmfSegment> RoundedPath(double left, double top, double right, double bottom, double rx, double ry)
        {
            double sx = right >= left ? 1d : -1d, sy = bottom >= top ? 1d : -1d;
            rx = Math.Abs(rx) * sx;
            ry = Math.Abs(ry) * sy;
            double kx = rx * Kappa, ky = ry * Kappa;
            var segments = new List<EmfSegment>
            {
                new(EmfSegmentKind.Move, left + rx, top, 0d, 0d, 0d, 0d),
                new(EmfSegmentKind.Line, right - rx, top, 0d, 0d, 0d, 0d),
                new(EmfSegmentKind.Curve, right - rx + kx, top, right, top + ry - ky, right, top + ry),
                new(EmfSegmentKind.Line, right, bottom - ry, 0d, 0d, 0d, 0d),
                new(EmfSegmentKind.Curve, right, bottom - ry + ky, right - rx + kx, bottom, right - rx, bottom),
                new(EmfSegmentKind.Line, left + rx, bottom, 0d, 0d, 0d, 0d),
                new(EmfSegmentKind.Curve, left + rx - kx, bottom, left, bottom - ry + ky, left, bottom - ry),
                new(EmfSegmentKind.Line, left, top + ry, 0d, 0d, 0d, 0d),
                new(EmfSegmentKind.Curve, left, top + ry - ky, left + rx - kx, top, left + rx, top),
                new(EmfSegmentKind.Close, 0d, 0d, 0d, 0d, 0d, 0d)
            };
            return segments;
        }

        private void PaintPath(bool fill, bool stroke)
        {
            if (finishedPath is null)
            {
                return;
            }

            List<EmfSegment> painted = finishedPath;
            finishedPath = null;
            Emit(painted, fill, stroke);
        }

        private void Emit(List<EmfSegment> segments, bool fill, bool stroke)
        {
            uint? fillColor = fill ? CurrentFill() : null;
            EmfStroke? outline = stroke ? CurrentStroke() : null;
            if (fillColor is null && outline is null || segments.Count == 0)
            {
                return;
            }

            ops.Add(new EmfPathOp { Segments = segments, EvenOdd = dc.PolyFillMode == 1, FillColor = fillColor, Stroke = outline, Clip = dc.Clip });
        }

        private void SelectClipPath(int mode)
        {
            if (finishedPath is null)
            {
                return;
            }

            List<EmfSegment> clipPath = finishedPath;
            finishedPath = null;
            if (mode is not (1 or 5))
            {
                throw new NotSupportedException("EMF clip path combination mode is not supported.");
            }

            if (mode == 5)
            {
                dc.Clip = null;
            }

            PushClip(clipPath, dc.PolyFillMode == 1);
        }

        private void PushClip(List<EmfSegment> deviceSegments, bool evenOdd)
        {
            if (dc.Clip is { Depth: >= 64 })
            {
                throw new NotSupportedException("EMF clip nesting is too deep.");
            }

            dc.Clip = new EmfClip(dc.Clip, deviceSegments, evenOdd);
        }

        private List<EmfSegment> MapToDevice(List<EmfSegment> logicalShape)
        {
            var device = new List<EmfSegment>(logicalShape.Count);
            foreach (EmfSegment segment in logicalShape)
            {
                (double x1, double y1) = ToDevice(segment.X1, segment.Y1);
                (double x2, double y2) = ToDevice(segment.X2, segment.Y2);
                (double x3, double y3) = ToDevice(segment.X3, segment.Y3);
                device.Add(segment with { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, X3 = x3, Y3 = y3 });
            }

            return device;
        }

        private void ExtSelectClipRegion(int p)
        {
            int regionBytes = I32(p + 8);
            int mode = I32(p + 12);
            if (mode == 5)
            {
                dc.Clip = null;
                if (regionBytes <= 0)
                {
                    return;
                }
            }
            else if (regionBytes <= 0)
            {
                // An empty operand: AND with nothing selects the full device, OR/XOR/DIFF are no-ops.
                return;
            }

            if (mode is not (1 or 5))
            {
                throw new NotSupportedException("EMF region combination mode is not supported.");
            }

            int rectangles = I32(p + 16 + 8);
            if (rectangles < 0 || regionBytes < 32 + rectangles * 16L || p + 16 + 32 + rectangles * 16L > data.Length)
            {
                throw new InvalidDataException("EMF clip region is truncated.");
            }

            var segments = new List<EmfSegment>();
            for (int i = 0; i < rectangles; i++)
            {
                int r = p + 16 + 32 + i * 16;
                // Region rectangles are device pixels and already mapped.
                segments.AddRange(RectanglePath(I32(r), I32(r + 4), I32(r + 8), I32(r + 12)));
            }

            PushClip(segments, evenOdd: false);
        }

        private void StretchDibits(int p)
        {
            if (I32(p + 68) != 0x00CC0020)
            {
                throw new NotSupportedException("EMF bitmap raster operation is not supported.");
            }

            AddBitmap(I32(p + 24), I32(p + 28), I32(p + 72), I32(p + 76), I32(p + 32), I32(p + 36), I32(p + 40), I32(p + 44),
                I32(p + 48), I32(p + 52), I32(p + 56), I32(p + 60), I32(p + 64), p);
        }

        private void BitBlt(int p, bool stretch)
        {
            int rop = I32(p + 40);
            int x = I32(p + 24), y = I32(p + 28), width = I32(p + 32), height = I32(p + 36);
            int infoLength = I32(p + 88);
            if (infoLength == 0)
            {
                // No source bitmap: solid fills only.
                uint? color = rop switch
                {
                    0x00F00021 => CurrentFill(),
                    0x00000042 => 0u,
                    0x00FF0062 => 0x00FFFFFFu,
                    _ => throw new NotSupportedException("EMF bit-block transfer operation is not supported.")
                };
                if (color is { } fillColor)
                {
                    ops.Add(new EmfPathOp { Segments = MapToDevice(RectanglePath(x, y, x + width, y + height)), FillColor = fillColor, Clip = dc.Clip });
                }

                return;
            }

            if (rop != 0x00CC0020)
            {
                throw new NotSupportedException("EMF bitmap raster operation is not supported.");
            }

            int sourceWidth = stretch ? I32(p + 100) : width;
            int sourceHeight = stretch ? I32(p + 104) : height;
            AddBitmap(x, y, width, height, I32(p + 44), I32(p + 48), sourceWidth, sourceHeight,
                I32(p + 84), infoLength, I32(p + 92), I32(p + 96), I32(p + 80), p);
        }

        private void AddBitmap(int destX, int destY, int destWidth, int destHeight, int srcX, int srcY, int srcWidth, int srcHeight,
            int infoOffsetInRecord, int infoLength, int bitsOffsetInRecord, int bitsLength, int usage, int recordStart)
        {
            if (usage != 0)
            {
                throw new NotSupportedException("EMF bitmap uses palette-index colors.");
            }

            if (infoLength <= 0 || infoOffsetInRecord <= 0 || bitsOffsetInRecord <= 0 || bitsLength <= 0)
            {
                throw new NotSupportedException("EMF bitmap record carries no source bitmap.");
            }

            int infoOffset = recordStart + infoOffsetInRecord;
            int bitsOffset = recordStart + bitsOffsetInRecord;
            if (infoOffset < 0 || infoOffset + (long)infoLength > data.Length || bitsOffset < 0 || bitsOffset + (long)bitsLength > data.Length)
            {
                throw new InvalidDataException("EMF bitmap data extends beyond the end of the file.");
            }

            (int bitmapWidth, int bitmapHeight) = DibImage.ReadSize(data.AsSpan(infoOffset, infoLength));
            bool bottomUp = I32(infoOffset + 8) > 0;
            int compression = I32(infoOffset + 16);
            if (compression is 1 or 2)
            {
                throw new NotSupportedException("EMF run-length encoded bitmaps are not supported.");
            }

            if (destWidth == 0 || destHeight == 0 || srcWidth <= 0 || srcHeight <= 0)
            {
                return;
            }

            // Source rectangles of bottom-up bitmaps count rows from the bottom edge.
            int topDownY = bottomUp ? bitmapHeight - (srcY + srcHeight) : srcY;
            int x0 = Math.Max(srcX, 0), x1 = Math.Min(srcX + srcWidth, bitmapWidth);
            int y0 = Math.Max(topDownY, 0), y1 = Math.Min(topDownY + srcHeight, bitmapHeight);
            if (x1 <= x0 || y1 <= y0)
            {
                return;
            }

            // Destination sub-rectangle that the clipped source maps to.
            double fx0 = (x0 - srcX) / (double)srcWidth, fx1 = (x1 - srcX) / (double)srcWidth;
            double fy0 = (y0 - topDownY) / (double)srcHeight, fy1 = (y1 - topDownY) / (double)srcHeight;
            var unitToLogical = new EmfMatrix(destWidth * (fx1 - fx0), 0d, 0d, destHeight * (fy1 - fy0), destX + destWidth * fx0, destY + destHeight * fy0);
            ops.Add(new EmfBitmapOp
            {
                UnitToDevice = unitToLogical.Then(LogicalToDevice()),
                InfoOffset = infoOffset,
                InfoLength = infoLength,
                BitsOffset = bitsOffset,
                BitsLength = bitsLength,
                SrcX = x0,
                SrcY = y0,
                SrcWidth = x1 - x0,
                SrcHeight = y1 - y0,
                Clip = dc.Clip
            });
        }

        private int I32(int offset) => BitConverter.ToInt32(data, offset);

        private int U32Int(int offset)
        {
            uint value = BitConverter.ToUInt32(data, offset);
            return value > int.MaxValue ? int.MaxValue : (int)value;
        }

        private double F32(int offset) => BitConverter.ToSingle(data, offset);
    }
}
