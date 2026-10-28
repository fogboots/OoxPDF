namespace Lokad.OoxPdf.Imaging;

internal sealed partial class EmfScene
{
    private sealed record WmfRegion;

    private sealed record WmfPalette;

    private sealed partial class EmfPlayer
    {
        // WMF objects take the lowest free slot of the object table when created.
        private readonly List<object?> wmfObjects = [];

        // Windows Metafile playback (MS-WMF). A metafile picture is played into a
        // device context whose viewport is the picture frame, so the device space here
        // is the metafile's first window (or the placeable header's bounds), and later
        // window records rescale the logical coordinates into it. Records carry 16-bit
        // parameters, mostly in reverse order (y before x).
        public EmfScene PlayWmf()
        {
            int position = 0;
            int inch = 0;
            (int Left, int Top, int Right, int Bottom)? bounds = null;
            if (U32(data, 0) == WmfPlaceableKey)
            {
                bounds = (S16(6), S16(8), S16(10), S16(12));
                inch = U16(14);
                position = 22;
            }

            position += U16(position + 2) * 2;

            // The first window origin and extent, which may follow the first drawing records.
            int? originX = null, originY = null, extentX = null, extentY = null;
            for (int p = position; p + 6 <= data.Length;)
            {
                long words = U32(data, p);
                int function = U16(p + 4);
                if (words < 3 || p + words * 2 > data.Length || function == 0)
                {
                    break;
                }

                if (function == 0x020B && originX is null && words >= 5)
                {
                    (originY, originX) = (S16(p + 6), S16(p + 8));
                }
                else if (function == 0x020C && extentX is null && words >= 5)
                {
                    (extentY, extentX) = (S16(p + 6), S16(p + 8));
                }

                p += (int)(words * 2);
            }

            if (extentX is null && bounds is { } b)
            {
                (originX, originY, extentX, extentY) = (b.Left, b.Top, b.Right - b.Left, b.Bottom - b.Top);
            }

            if (extentX is not { } windowWidth || extentY is not { } windowHeight || windowWidth == 0 || windowHeight == 0)
            {
                throw new NotSupportedException("WMF has neither a window extent nor placeable bounds.");
            }

            frameLeft = originX ?? 0;
            frameTop = originY ?? 0;
            frameWidth = Math.Abs(windowWidth);
            frameHeight = Math.Abs(windowHeight);
            dc.MapMode = 8;
            (dc.WindowOrgX, dc.WindowOrgY, dc.WindowExtX, dc.WindowExtY) = (frameLeft, frameTop, windowWidth, windowHeight);
            (dc.ViewportOrgX, dc.ViewportOrgY, dc.ViewportExtX, dc.ViewportExtY) = (frameLeft, frameTop, frameWidth, frameHeight);

            // Cosmetic pens are one output pixel (1/96 inch) wide; a picture without a
            // placeable header is assumed to be in twips.
            double unitsPerInch = inch > 0 ? inch : 1440d;
            if (inch > 0 && bounds is { } placed && placed.Right != placed.Left)
            {
                unitsPerInch *= frameWidth / Math.Abs(placed.Right - placed.Left);
            }

            cosmeticPenWidth = unitsPerInch / 96d;
            minimumPenWidth = 0d;

            int count = 0;
            while (position + 6 <= data.Length)
            {
                if ((++count & 1023) == 0)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                }

                long words = U32(data, position);
                int function = U16(position + 4);
                if (words < 3 || position + words * 2 > data.Length)
                {
                    throw new InvalidDataException("WMF record extends beyond the end of the data.");
                }

                if (function == 0)
                {
                    break;
                }

                PlayWmfRecord(function, position, (int)(words * 2));
                position += (int)(words * 2);
            }

            return new EmfScene(data, frameLeft, frameTop, frameWidth, frameHeight, ops);
        }

        private void PlayWmfRecord(int function, int p, int size)
        {
            int q = p + 6;
            switch (function)
            {
                case 0x020B: Need(size, 4); dc.WindowOrgX = S16(q + 2); dc.WindowOrgY = S16(q); break;
                case 0x020C: Need(size, 4); dc.WindowExtX = S16(q + 2); dc.WindowExtY = S16(q); break;
                case 0x020F: Need(size, 4); dc.WindowOrgX += S16(q + 2); dc.WindowOrgY += S16(q); break; // META_OFFSETWINDOWORG
                case 0x0410: // META_SCALEWINDOWEXT: yDenom, yNum, xDenom, xNum
                    Need(size, 8);
                    dc.WindowExtX = Scale(dc.WindowExtX, S16(q + 6), S16(q + 4));
                    dc.WindowExtY = Scale(dc.WindowExtY, S16(q + 2), S16(q));
                    break;
                case 0x0103: // META_SETMAPMODE: pictures play in the caller's anisotropic mapping
                    Need(size, 2);
                    if (U16(q) is not (1 or 7 or 8))
                    {
                        throw new NotSupportedException("WMF mapping mode is not supported.");
                    }

                    break;
                case 0x0102: Need(size, 2); dc.BkMode = U16(q); break;
                case 0x0201: Need(size, 4); dc.BkColor = U32(data, q); break;
                case 0x0209: Need(size, 4); dc.TextColor = U32(data, q); break;
                case 0x012E: Need(size, 2); dc.TextAlign = U16(q); break;
                case 0x0104: Need(size, 2); dc.Rop2 = U16(q); break;
                case 0x0106: Need(size, 2); dc.PolyFillMode = U16(q); break;
                case 0x001E: savedContexts.Add(dc.Clone()); break;
                case 0x0127: Need(size, 2); RestoreContext(S16(q)); break;
                case 0x0214: Need(size, 4); currentX = S16(q + 2); currentY = S16(q); break;
                case 0x0213: Need(size, 4); LineTo(S16(q + 2), S16(q)); break;
                case 0x0324: WmfPoly(q, size, close: true); break;
                case 0x0325: WmfPoly(q, size, close: false); break;
                case 0x0538: WmfPolyPolygon(q, size); break;
                case 0x041B: Need(size, 8); Shape(RectanglePath(S16(q + 6), S16(q + 4), S16(q + 2), S16(q))); break;
                case 0x0418: Need(size, 8); Shape(EllipsePath(S16(q + 6), S16(q + 4), S16(q + 2), S16(q))); break;
                case 0x061C: Need(size, 12); Shape(RoundRectanglePath(S16(q + 10), S16(q + 8), S16(q + 6), S16(q + 4), S16(q + 2), S16(q))); break;
                case 0x0416: // META_INTERSECTCLIPRECT: bottom, right, top, left
                    Need(size, 8);
                    PushClip(MapToDevice(RectanglePath(S16(q + 6), S16(q + 4), S16(q + 2), S16(q))), evenOdd: false);
                    break;
                case 0x0415: Need(size, 8); ExcludeClipRectangle(S16(q + 6), S16(q + 4), S16(q + 2), S16(q)); break;
                case 0x012C: // META_SELECTCLIPREGION: an index that names no region selects none
                    Need(size, 2);
                    if (WmfObject(U16(q)) is WmfRegion)
                    {
                        throw new NotSupportedException("WMF clip regions are not supported.");
                    }

                    dc.Clip = null;
                    break;
                case 0x02FA: WmfCreatePen(q, size); break;
                case 0x02FC: WmfCreateBrush(q, size); break;
                case 0x02FB: WmfCreateFont(q, size); break;
                case 0x00F7: AddWmfObject(new WmfPalette()); break;
                case 0x06FF: AddWmfObject(new WmfRegion()); break;
                case 0x01F9: case 0x0142: AddWmfObject(new BrushObject(false, 0u, "pattern brushes")); break;
                case 0x012D: // META_SELECTOBJECT
                    Need(size, 2);
                    switch (WmfObject(U16(q)))
                    {
                        case WmfRegion: throw new NotSupportedException("WMF clip regions are not supported.");
                        case { } selected: Select(selected); break;
                    }

                    break;
                case 0x01F0: // META_DELETEOBJECT
                    Need(size, 2);
                    if (U16(q) < wmfObjects.Count)
                    {
                        wmfObjects[U16(q)] = null;
                    }

                    break;
                case 0x0A32: WmfExtTextOut(q, size); break;
                case 0x0521: WmfTextOut(q, size); break;
                case 0x0940: WmfDibBitBlt(p, q, size, stretch: false); break;
                case 0x0B41: WmfDibBitBlt(p, q, size, stretch: true); break;
                case 0x0F43: WmfStretchDib(p, q, size); break;
                case 0x061D: // META_PATBLT: raster operation, height, width, y, x
                    Need(size, 12);
                    FillBlock(U32(data, q), S16(q + 10), S16(q + 8), S16(q + 6), S16(q + 4));
                    break;
                case 0x0626: case 0x0107: case 0x0231: case 0x0105: case 0x0108: case 0x020A: case 0x0149:
                case 0x0234: case 0x0035: case 0x0037: case 0x0139: case 0x0436:
                    // Escapes (comments, equation data), stretch mode, font mapper flags, character
                    // extra and justification (the drawn advances already include them), layout and
                    // palette plumbing: none changes what this player paints.
                    break;
                default:
                    throw new NotSupportedException($"WMF record function 0x{function:X4} is not supported.");
            }
        }

        private static void Need(int size, int parameterBytes)
        {
            if (size < 6 + parameterBytes)
            {
                throw new InvalidDataException("WMF record is truncated.");
            }
        }

        private object? WmfObject(int index)
        {
            return index < wmfObjects.Count ? wmfObjects[index] : null;
        }

        private void AddWmfObject(object created)
        {
            int free = wmfObjects.IndexOf(null);
            if (free < 0)
            {
                wmfObjects.Add(created);
            }
            else
            {
                wmfObjects[free] = created;
            }
        }

        private void WmfCreatePen(int q, int size)
        {
            // Style, width (x; y is unused), color.
            Need(size, 10);
            int style = U16(q);
            int type = style & 0xF;
            int cap = (style & 0xF00) switch { 0x100 => 2, 0x200 => 0, _ => 1 };
            int join = (style & 0xF000) switch { 0x1000 => 2, 0x2000 => 0, _ => 1 };
            AddWmfObject(new PenObject(type == 5, U32(data, q + 6), S16(q + 2), type, cap, join, false, null));
        }

        private void WmfCreateBrush(int q, int size)
        {
            Need(size, 8);
            int style = U16(q);
            string? unsupported = style switch { 0 or 1 => null, 2 => "hatched brushes", _ => "pattern brushes" };
            AddWmfObject(new BrushObject(style == 1, U32(data, q + 2), unsupported));
        }

        private void WmfCreateFont(int q, int size)
        {
            // Height, width, escapement, orientation, weight, italic, underline, strike-out,
            // charset, precision and quality bytes, then the face name (at most 32 bytes).
            Need(size, 18);
            int nameLength = Math.Min(32, size - 6 - 18);
            ReadOnlySpan<byte> name = data.AsSpan(q + 18, nameLength);
            int end = name.IndexOf((byte)0);
            string face = System.Text.Encoding.Latin1.GetString(end < 0 ? name : name[..end]);
            AddWmfObject(new FontObject(face, S16(q), S16(q + 4), S16(q + 8), data[q + 10] != 0, data[q + 11] != 0, data[q + 12] != 0, data[q + 13]));
        }

        private void WmfPoly(int q, int size, bool close)
        {
            Need(size, 2);
            int count = S16(q);
            if (count < 0 || 6 + 2 + count * 4L > size)
            {
                throw new InvalidDataException("WMF polygon record is truncated.");
            }

            if (count == 0)
            {
                return;
            }

            var logical = new List<EmfSegment>(count + 1);
            for (int i = 0; i < count; i++)
            {
                logical.Add(new EmfSegment(i == 0 ? EmfSegmentKind.Move : EmfSegmentKind.Line, S16(q + 2 + i * 4), S16(q + 4 + i * 4), 0d, 0d, 0d, 0d));
            }

            if (close)
            {
                logical.Add(new EmfSegment(EmfSegmentKind.Close, 0d, 0d, 0d, 0d, 0d, 0d));
            }

            Emit(MapToDevice(logical), fill: close, stroke: true);
        }

        private void WmfPolyPolygon(int q, int size)
        {
            Need(size, 2);
            int polygons = S16(q);
            if (polygons < 0 || 6 + 2 + polygons * 2L > size)
            {
                throw new InvalidDataException("WMF poly-polygon record is truncated.");
            }

            var logical = new List<EmfSegment>();
            int point = q + 2 + polygons * 2;
            for (int polygon = 0; polygon < polygons; polygon++)
            {
                int count = S16(q + 2 + polygon * 2);
                if (count < 0 || point - (q - 6) + count * 4L > size)
                {
                    throw new InvalidDataException("WMF poly-polygon counts are inconsistent.");
                }

                for (int i = 0; i < count; i++, point += 4)
                {
                    logical.Add(new EmfSegment(i == 0 ? EmfSegmentKind.Move : EmfSegmentKind.Line, S16(point), S16(point + 2), 0d, 0d, 0d, 0d));
                }

                if (count > 0)
                {
                    logical.Add(new EmfSegment(EmfSegmentKind.Close, 0d, 0d, 0d, 0d, 0d, 0d));
                }
            }

            Emit(MapToDevice(logical), fill: true, stroke: true);
        }

        private void WmfExtTextOut(int q, int size)
        {
            // y, x, string length, options, [rectangle: left, top, right, bottom], string
            // (padded to a word), optional advances.
            Need(size, 8);
            int y = S16(q), x = S16(q + 2), count = S16(q + 4), options = U16(q + 6);
            int text = q + 8;
            int[]? rectangle = null;
            if ((options & (EtoOpaque | EtoClipped)) != 0)
            {
                Need(size, 16);
                rectangle = [S16(q + 8), S16(q + 10), S16(q + 12), S16(q + 14)];
                text += 8;
            }

            int padded = count + (count & 1);
            if (count < 0 || text - (q - 6) + (long)count > size)
            {
                throw new InvalidDataException("WMF text record is truncated.");
            }

            double[]? advances = null;
            int dx = text + padded;
            if (count > 0 && dx - (q - 6) + count * 2L <= size)
            {
                advances = new double[count];
                for (int i = 0; i < count; i++)
                {
                    advances[i] = S16(dx + i * 2);
                }
            }

            DrawWmfText(x, y, data.AsSpan(text, count), advances, options, rectangle);
        }

        private void WmfTextOut(int q, int size)
        {
            // String length, string (padded to a word), y, x.
            Need(size, 2);
            int count = S16(q);
            int padded = count + (count & 1);
            if (count < 0 || 6 + 2 + padded + 4L > size)
            {
                throw new InvalidDataException("WMF text record is truncated.");
            }

            DrawWmfText(S16(q + 2 + padded + 2), S16(q + 2 + padded), data.AsSpan(q + 2, count), null, 0, null);
        }

        private void DrawWmfText(int x, int y, ReadOnlySpan<byte> text, double[]? advances, int options, int[]? rectangle)
        {
            if (dc.Font is not { } font)
            {
                throw new NotSupportedException("WMF text uses the stock device font.");
            }

            DrawText(x, y, MetafileCharacters.Decode(font.FaceName, font.Charset, text), advances, options, rectangle, advancedGraphics: false);
        }

        private void WmfDibBitBlt(int p, int q, int size, bool stretch)
        {
            // Without a bitmap the record is (function >> 8) + 3 words and paints with the brush.
            int function = U16(p + 4);
            uint rop = U32(data, q);
            if (size == ((function >> 8) + 3) * 2)
            {
                // Raster operation, [source height, width,] source y, x, reserved, height, width, y, x.
                int o = stretch ? q + 14 : q + 10;
                FillBlock(rop, S16(o + 6), S16(o + 4), S16(o + 2), S16(o));
                return;
            }

            if (stretch)
            {
                // Raster operation, source height, width, y, x, destination height, width, y, x, DIB.
                Need(size, 22);
                WmfBitmap(p, q + 20, size, rop, S16(q + 18), S16(q + 16), S16(q + 14), S16(q + 12), S16(q + 10), S16(q + 8), S16(q + 6), S16(q + 4), 0);
            }
            else
            {
                // Raster operation, source y, x, height, width, destination y, x, DIB.
                Need(size, 18);
                int height = S16(q + 8), width = S16(q + 10);
                WmfBitmap(p, q + 16, size, rop, S16(q + 14), S16(q + 12), width, height, S16(q + 6), S16(q + 4), width, height, 0);
            }
        }

        private void WmfStretchDib(int p, int q, int size)
        {
            // Raster operation, color usage, source height, width, y, x, destination height, width, y, x, DIB.
            Need(size, 24);
            WmfBitmap(p, q + 22, size, U32(data, q), S16(q + 20), S16(q + 18), S16(q + 16), S16(q + 14), S16(q + 12), S16(q + 10), S16(q + 8), S16(q + 6), U16(q + 4));
        }

        private void WmfBitmap(int p, int dib, int size, uint rop, int destX, int destY, int destWidth, int destHeight,
            int srcX, int srcY, int srcWidth, int srcHeight, int usage)
        {
            if (rop == 0x00AA0029u)
            {
                return; // the destination is left unchanged
            }

            if (rop != 0x00CC0020u)
            {
                throw new NotSupportedException("WMF bitmap raster operation is not supported.");
            }

            int available = p + size - dib;
            if (available < 40)
            {
                throw new InvalidDataException("WMF bitmap is truncated.");
            }

            int headerSize = I32(dib);
            if (headerSize < 40 || headerSize > available)
            {
                throw new NotSupportedException("WMF bitmap header is not supported.");
            }

            int bitCount = U16(dib + 14), compression = I32(dib + 16), used = I32(dib + 32);
            int colors = used > 0 ? used : bitCount <= 8 ? 1 << bitCount : 0;
            int masks = headerSize == 40 ? compression switch { 3 => 12, 6 => 16, _ => 0 } : 0;
            long infoLength = headerSize + masks + colors * 4L;
            if (colors < 0 || infoLength >= available)
            {
                throw new InvalidDataException("WMF bitmap color table is truncated.");
            }

            AddBitmap(destX, destY, destWidth, destHeight, srcX, srcY, srcWidth, srcHeight,
                dib - p, (int)infoLength, dib - p + (int)infoLength, available - (int)infoLength, usage, p);
        }

        // A bit-block transfer without a source bitmap: brush, black or white fills only.
        private void FillBlock(uint rop, int x, int y, int width, int height)
        {
            uint? color = rop switch
            {
                0x00F00021u => CurrentFill(),
                0x00000042u => 0u,
                0x00FF0062u => 0x00FFFFFFu,
                0x00AA0029u => null, // the destination is left unchanged
                _ => throw new NotSupportedException("WMF bit-block transfer operation is not supported.")
            };
            if (color is { } fillColor)
            {
                ops.Add(new EmfPathOp { Segments = MapToDevice(RectanglePath(x, y, x + width, y + height)), FillColor = fillColor, Clip = dc.Clip });
            }
        }

        private int S16(int offset) => (short)(data[offset] | (data[offset + 1] << 8));

        private int U16(int offset) => data[offset] | (data[offset + 1] << 8);
    }
}
