using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Imaging;

// The face that draws a code point for a requested family and style (the family
// itself or a fallback that has the glyph), and the styles it must synthesize.
internal readonly record struct EmfGlyphFont(OpenTypeFont Font, bool SyntheticBold, bool SyntheticItalic);

internal delegate EmfGlyphFont? EmfFontLookup(string familyName, bool bold, bool italic, int codePoint);

internal sealed partial class EmfScene
{
    private sealed record FontObject(string FaceName, int Height, int Escapement, int Weight, bool Italic, bool Underline, bool Strikeout, int Charset);

    private sealed partial class EmfPlayer
    {
        private const int TaUpdateCp = 1, TaRight = 2, TaCenter = 6, TaBottom = 8, TaBaseline = 24;
        private const int EtoOpaque = 2, EtoClipped = 4, EtoGlyphIndex = 0x10, EtoRtlReading = 0x80, EtoPdy = 0x2000;

        private readonly Dictionary<(string Family, bool Bold, bool Italic, int CodePoint), EmfGlyphFont?> glyphFonts = [];

        private void CreateFont(int p, int size)
        {
            // EMR_EXTCREATEFONTINDIRECTW: object index, then LOGFONTW.
            if (size < 104)
            {
                throw new InvalidDataException("EMF font record is truncated.");
            }

            string face = new string(System.Runtime.InteropServices.MemoryMarshal.Cast<byte, char>(data.AsSpan(p + 40, 64))).Split('\0')[0];
            objects[I32(p + 8)] = new FontObject(face, I32(p + 12), I32(p + 20), I32(p + 28), data[p + 32] != 0, data[p + 33] != 0, data[p + 34] != 0, data[p + 35]);
        }

        // EMR_EXTTEXTOUTA/W: bounds, graphics mode, scales, then EMRTEXT.
        private void ExtTextOut(int p, int size, bool wide)
        {
            int count = I32(p + 44), stringOffset = I32(p + 48), options = I32(p + 52), dxOffset = I32(p + 72);
            int unit = wide ? 2 : 1;
            if (count < 0 || stringOffset < 0 || (long)stringOffset + (long)count * unit > size)
            {
                throw new InvalidDataException("EMF text record is truncated.");
            }

            if (dc.Font is not { } font)
            {
                throw new NotSupportedException("EMF text uses the stock device font.");
            }

            int[] codePoints = wide
                ? MetafileCharacters.Decode(font.FaceName, font.Charset, System.Runtime.InteropServices.MemoryMarshal.Cast<byte, char>(data.AsSpan(p + stringOffset, count * 2)))
                : MetafileCharacters.Decode(font.FaceName, font.Charset, data.AsSpan(p + stringOffset, count));
            double[]? advances = null;
            if (dxOffset > 0 && (long)dxOffset + (long)count * 4 <= size)
            {
                advances = new double[count];
                for (int i = 0; i < count; i++)
                {
                    advances[i] = I32(p + dxOffset + i * 4);
                }
            }

            int[]? rectangle = (options & (EtoOpaque | EtoClipped)) != 0 ? [I32(p + 56), I32(p + 60), I32(p + 64), I32(p + 68)] : null;
            DrawText(I32(p + 36), I32(p + 40), codePoints, advances, options, rectangle, advancedGraphics: I32(p + 24) == 2);
        }

        private void ExcludeClipRectangle(double left, double top, double right, double bottom)
        {
            // The frame enlarged on every side stands for the whole device; the even-odd
            // rule then removes the rectangle from it.
            List<EmfSegment> segments = RectanglePath(frameLeft - frameWidth, frameTop - frameHeight, frameLeft + 2d * frameWidth, frameTop + 2d * frameHeight);
            segments.AddRange(MapToDevice(RectanglePath(left, top, right, bottom)));
            PushClip(segments, evenOdd: true);
        }

        // Draws one text record at reference point (x, y) in logical units. codePoints
        // are Unicode (below 0x20: advance only, -1: the low half of a surrogate pair);
        // advances are the logical distances between character origins, or null for the
        // font's own advance widths. In the compatible graphics mode GDI draws text
        // upright in device space at the mapped size; the advanced mode transforms it
        // with the logical coordinates.
        private void DrawText(double x, double y, int[] codePoints, double[]? advances, int options, int[]? rectangle, bool advancedGraphics)
        {
            if ((options & (EtoGlyphIndex | EtoPdy | EtoRtlReading)) != 0)
            {
                throw new NotSupportedException("Metafile text with glyph indices, vertical offsets or right-to-left reading is not supported.");
            }

            if (dc.Font is not { } font)
            {
                throw new NotSupportedException("Metafile text uses the stock device font.");
            }

            if (fonts is null)
            {
                throw new NotSupportedException("Metafile text records are not supported without fonts.");
            }

            if (font.Height == 0)
            {
                throw new NotSupportedException("Metafile text uses the default font height.");
            }

            if (rectangle is not null && (options & EtoOpaque) != 0)
            {
                ops.Add(new EmfPathOp { Segments = MapToDevice(RectanglePath(rectangle[0], rectangle[1], rectangle[2], rectangle[3])), FillColor = TextStateColor(dc.BkColor), Clip = dc.Clip });
            }

            if (codePoints.Length == 0)
            {
                return;
            }

            string family = MetafileCharacters.DrawingFamily(font.FaceName);
            bool bold = font.Weight >= 600, italic = font.Italic, vertical = MetafileCharacters.IsVertical(font.FaceName);
            var glyphs = new (EmfGlyphFont Source, ushort GlyphId)?[codePoints.Length];
            var upright = new bool[codePoints.Length];
            for (int i = 0; i < codePoints.Length; i++)
            {
                int codePoint = codePoints[i];
                if (codePoint < 0x20)
                {
                    continue;
                }

                EmfGlyphFont source = Lookup(family, bold, italic, codePoint) ??
                    throw new NotSupportedException($"No installed font draws U+{codePoint:X4} of metafile font '{font.FaceName}'.");
                if (!source.Font.HasTrueTypeOutlines)
                {
                    throw new NotSupportedException($"The font drawing U+{codePoint:X4} of metafile font '{font.FaceName}' has no TrueType outlines.");
                }

                ushort glyphId = source.Font.MapCodePoint(codePoint);
                if (vertical && MetafileCharacters.IsUprightInVertical(codePoint, out bool needsAlternate))
                {
                    ushort alternate = source.Font.VerticalAlternate(glyphId);
                    upright[i] = !needsAlternate || alternate != glyphId;
                    glyphId = alternate;
                }

                glyphs[i] = (source, glyphId);
            }

            OpenTypeFont metrics = (Lookup(family, bold, italic, 'x') ?? glyphs.FirstOrDefault(glyph => glyph is not null)?.Source ??
                throw new NotSupportedException("Metafile text has no drawable characters.")).Font;
            double unitsPerEm = metrics.UnitsPerEm;
            int cellUnits = metrics.Os2.WindowsAscender + metrics.Os2.WindowsDescender;
            if (unitsPerEm <= 0d || cellUnits <= 0)
            {
                throw new NotSupportedException("Metafile text font has no usable vertical metrics.");
            }

            if ((dc.TextAlign & TaUpdateCp) != 0)
            {
                x = currentX;
                y = currentY;
            }

            // Text space: device pixels in the compatible mode, logical units in the advanced mode.
            EmfMatrix toDevice = LogicalToDevice();
            double unitX = 1d, unitY = 1d, originX = x, originY = y;
            if (!advancedGraphics)
            {
                unitX = Math.Sqrt(toDevice.A * toDevice.A + toDevice.B * toDevice.B);
                unitY = Math.Sqrt(toDevice.C * toDevice.C + toDevice.D * toDevice.D);
                (originX, originY) = toDevice.Apply(x, y);
            }

            // A negative height is the em size, a positive one the cell height.
            double em = Math.Abs(font.Height) * unitY * (font.Height > 0 ? unitsPerEm / cellUnits : 1d);
            double stretch = unitY > 0d ? unitX / unitY : 1d;
            double angle = font.Escapement / 10d * Math.PI / 180d;
            double ux = Math.Cos(angle), uy = -Math.Sin(angle); // along the baseline
            double vx = -Math.Sin(angle), vy = -Math.Cos(angle); // towards the ascender

            var steps = new double[codePoints.Length];
            double width = 0d;
            for (int i = 0; i < codePoints.Length; i++)
            {
                steps[i] = advances is not null
                    ? advances[i] * unitX
                    : glyphs[i] is { } glyph ? glyph.Source.Font.GetAdvanceWidth(glyph.GlyphId) * em / glyph.Source.Font.UnitsPerEm * stretch : 0d;
                width += steps[i];
            }

            int align = dc.TextAlign;
            double shift = (align & TaCenter) == TaCenter ? -width / 2d : (align & TaRight) != 0 ? -width : 0d;
            double ascent = em * metrics.Os2.WindowsAscender / unitsPerEm, descent = em * metrics.Os2.WindowsDescender / unitsPerEm;
            double rise = (align & TaBaseline) == TaBaseline ? 0d : (align & TaBottom) != 0 ? descent : -ascent;
            double baseX = originX + shift * ux + rise * vx, baseY = originY + shift * uy + rise * vy;

            if ((align & TaUpdateCp) != 0 && (align & TaCenter) != TaCenter)
            {
                double move = (align & TaRight) != 0 ? -width : width;
                if (advancedGraphics)
                {
                    currentX += move * ux;
                    currentY += move * uy;
                }
                else
                {
                    currentX += toDevice.A != 0d ? move * ux / toDevice.A : 0d;
                    currentY += toDevice.D != 0d ? move * uy / toDevice.D : 0d;
                }
            }

            (double X, double Y) Map(double along, double up)
            {
                double tx = baseX + along * ux + up * vx, ty = baseY + along * uy + up * vy;
                return advancedGraphics ? toDevice.Apply(tx, ty) : (tx, ty);
            }

            List<EmfSegment> Quad(double left, double right, double bottom, double top)
            {
                (double x0, double y0) = Map(left, bottom);
                (double x1, double y1) = Map(right, bottom);
                (double x2, double y2) = Map(right, top);
                (double x3, double y3) = Map(left, top);
                return
                [
                    new(EmfSegmentKind.Move, x0, y0, 0d, 0d, 0d, 0d),
                    new(EmfSegmentKind.Line, x1, y1, 0d, 0d, 0d, 0d),
                    new(EmfSegmentKind.Line, x2, y2, 0d, 0d, 0d, 0d),
                    new(EmfSegmentKind.Line, x3, y3, 0d, 0d, 0d, 0d),
                    new(EmfSegmentKind.Close, 0d, 0d, 0d, 0d, 0d, 0d)
                ];
            }

            EmfClip? clip = dc.Clip;
            if (rectangle is not null && (options & EtoClipped) != 0)
            {
                clip = new EmfClip(clip, MapToDevice(RectanglePath(rectangle[0], rectangle[1], rectangle[2], rectangle[3])), evenOdd: false);
            }

            if (dc.BkMode == 2 && (options & EtoOpaque) == 0)
            {
                ops.Add(new EmfPathOp { Segments = Quad(0d, width, -descent, ascent), FillColor = TextStateColor(dc.BkColor), Clip = clip });
            }

            var segments = new List<EmfSegment>();
            bool syntheticBold = false;
            double pen = 0d;
            for (int i = 0; i < codePoints.Length; i++)
            {
                if (glyphs[i] is { } glyph && glyph.Source.Font.TryReadGlyphOutline(glyph.GlyphId, out OpenTypeFont.OpenTypeGlyphOutline outline))
                {
                    double scale = em / glyph.Source.Font.UnitsPerEm;
                    double shear = glyph.Source.SyntheticItalic ? PdfGraphicsBuilder.SyntheticItalicShear : 0d;
                    double start = pen;
                    if (upright[i])
                    {
                        // A quarter turn counterclockwise about the centre of the glyph's cell,
                        // so that the glyph stands upright once the line runs downwards.
                        double centreX = steps[i] / 2d, centreY = (ascent - descent) / 2d;
                        AppendGlyph(segments, outline, (gx, gy) =>
                        {
                            double along = (gx + gy * shear) * scale * stretch - centreX, up = gy * scale - centreY;
                            return Map(start + centreX - up, centreY + along);
                        });
                    }
                    else
                    {
                        AppendGlyph(segments, outline, (gx, gy) => Map(start + (gx + gy * shear) * scale * stretch, gy * scale));
                    }
                    syntheticBold |= glyph.Source.SyntheticBold;
                }

                pen += steps[i];
            }

            if (font.Underline)
            {
                double position = em * metrics.Post.UnderlinePosition / unitsPerEm;
                double thickness = Math.Max(em * metrics.Post.UnderlineThickness / unitsPerEm, em / 20d);
                segments.AddRange(Quad(0d, width, position - thickness / 2d, position + thickness / 2d));
            }

            if (font.Strikeout)
            {
                double position = em * metrics.Os2.StrikeoutPosition / unitsPerEm;
                double thickness = Math.Max(em * metrics.Os2.StrikeoutSize / unitsPerEm, em / 20d);
                segments.AddRange(Quad(0d, width, position, position + thickness));
            }

            if (segments.Count == 0)
            {
                return;
            }

            uint color = TextStateColor(dc.TextColor);
            double deviceScale = advancedGraphics ? toDevice.ScaleFactor : 1d;
            ops.Add(new EmfPathOp
            {
                Segments = segments,
                FillColor = color,
                Stroke = syntheticBold ? new EmfStroke(color, em * deviceScale / 35d, null, 1, 1, dc.MiterLimit) : null,
                Clip = clip
            });
        }

        private EmfGlyphFont? Lookup(string family, bool bold, bool italic, int codePoint)
        {
            var key = (family, bold, italic, codePoint);
            if (!glyphFonts.TryGetValue(key, out EmfGlyphFont? found))
            {
                found = fonts!(family, bold, italic, codePoint);
                glyphFonts[key] = found;
            }

            return found;
        }

        private static uint TextStateColor(uint color)
        {
            // PALETTERGB (0x02) carries an RGB value; a palette index has none.
            if ((color & 0xFF000000u) == 0x01000000u)
            {
                throw new NotSupportedException("Metafile text uses a palette color.");
            }

            return color & 0x00FFFFFFu;
        }

        // Appends a TrueType outline (quadratic contours in font units, y up) as closed
        // figures; map places a font-unit point in device space.
        private static void AppendGlyph(List<EmfSegment> target, OpenTypeFont.OpenTypeGlyphOutline outline, Func<double, double, (double X, double Y)> map)
        {
            foreach (OpenTypeFont.OpenTypeGlyphContour contour in outline.Contours)
            {
                IReadOnlyList<OpenTypeFont.OpenTypeGlyphPoint> points = contour.Points;
                int count = points.Count;
                if (count == 0)
                {
                    continue;
                }

                int first = -1;
                for (int i = 0; i < count && first < 0; i++)
                {
                    if (points[i].IsOnCurve)
                    {
                        first = i;
                    }
                }

                // A contour of off-curve points starts at the implied point between its ends.
                (double X, double Y) start = first >= 0
                    ? (points[first].X, points[first].Y)
                    : ((points[count - 1].X + points[0].X) / 2d, (points[count - 1].Y + points[0].Y) / 2d);
                (double X, double Y) current = start;
                (double X, double Y)? control = null;
                (double sx, double sy) = map(start.X, start.Y);
                target.Add(new EmfSegment(EmfSegmentKind.Move, sx, sy, 0d, 0d, 0d, 0d));
                int remaining = first >= 0 ? count - 1 : count;
                for (int j = 0; j < remaining; j++)
                {
                    OpenTypeFont.OpenTypeGlyphPoint point = points[(first + 1 + j) % count];
                    if (point.IsOnCurve)
                    {
                        if (control is { } c)
                        {
                            AddQuadratic(c, (point.X, point.Y));
                        }
                        else
                        {
                            (double lx, double ly) = map(point.X, point.Y);
                            target.Add(new EmfSegment(EmfSegmentKind.Line, lx, ly, 0d, 0d, 0d, 0d));
                        }

                        current = (point.X, point.Y);
                        control = null;
                    }
                    else
                    {
                        if (control is { } c)
                        {
                            var middle = ((c.X + point.X) / 2d, (c.Y + point.Y) / 2d);
                            AddQuadratic(c, middle);
                            current = middle;
                        }

                        control = (point.X, point.Y);
                    }
                }

                if (control is { } last)
                {
                    AddQuadratic(last, start);
                }

                target.Add(new EmfSegment(EmfSegmentKind.Close, 0d, 0d, 0d, 0d, 0d, 0d));

                void AddQuadratic((double X, double Y) c, (double X, double Y) end)
                {
                    (double x1, double y1) = map(current.X + (c.X - current.X) * 2d / 3d, current.Y + (c.Y - current.Y) * 2d / 3d);
                    (double x2, double y2) = map(end.X + (c.X - end.X) * 2d / 3d, end.Y + (c.Y - end.Y) * 2d / 3d);
                    (double x3, double y3) = map(end.X, end.Y);
                    target.Add(new EmfSegment(EmfSegmentKind.Curve, x1, y1, x2, y2, x3, y3));
                }
            }
        }
    }
}
