using System.Globalization;
using System.Text;

namespace Lokad.OoxPdf.Pdf;

internal sealed record PdfTilingPatternResource(string ResourceName, PdfTilingPattern Pattern);

internal sealed class PdfTilingPattern
{
    private string? resourceKey;

    public PdfTilingPattern(double width, double height, string content)
        : this(width, height, width, height, null, 1, content, [])
    {
    }

    public PdfTilingPattern(
        double width,
        double height,
        double xStep,
        double yStep,
        PdfPatternMatrix? matrix,
        int tilingType,
        string content)
        : this(width, height, xStep, yStep, matrix, tilingType, content, [])
    {
    }

    public PdfTilingPattern(
        double width,
        double height,
        double xStep,
        double yStep,
        PdfPatternMatrix? matrix,
        int tilingType,
        string content,
        IReadOnlyList<PdfImageResource> images)
    {
        Width = Math.Max(0.001d, width);
        Height = Math.Max(0.001d, height);
        XStep = Math.Max(0.001d, xStep);
        YStep = Math.Max(0.001d, yStep);
        Matrix = matrix;
        TilingType = Math.Clamp(tilingType, 1, 3);
        Content = content;
        Images = images;
    }

    public double Width { get; }

    public double Height { get; }

    public double XStep { get; }

    public double YStep { get; }

    public PdfPatternMatrix? Matrix { get; }

    public int TilingType { get; }

    public string Content { get; }

    public IReadOnlyList<PdfImageResource> Images { get; }

    public string ResourceKey => resourceKey ??= BuildResourceKey();

    public static PdfTilingPattern OfficeBitmapDiagonalLines(
        bool up,
        byte foregroundRed,
        byte foregroundGreen,
        byte foregroundBlue,
        byte backgroundRed,
        byte backgroundGreen,
        byte backgroundBlue)
    {
        const int cellSize = 16;
        const double matrixScale = 0.375d;
        const string imageName = "ImPattern";

        byte[] rgb = new byte[cellSize * cellSize * 3];
        for (int y = 0; y < cellSize; y++)
        {
            for (int x = 0; x < cellSize; x++)
            {
                int rowOffset = 2 * (y / 2);
                int diagonal = up
                    ? PositiveModulo(x + rowOffset, 8)
                    : PositiveModulo(x - rowOffset, 8);
                bool foreground = diagonal is 6 or 7;
                int index = (y * cellSize + x) * 3;
                rgb[index] = foreground ? foregroundRed : backgroundRed;
                rgb[index + 1] = foreground ? foregroundGreen : backgroundGreen;
                rgb[index + 2] = foreground ? foregroundBlue : backgroundBlue;
            }
        }

        var image = PdfImageXObject.RgbPng(cellSize, cellSize, rgb, alpha: null);
        return new PdfTilingPattern(
            cellSize,
            cellSize,
            cellSize,
            cellSize,
            new PdfPatternMatrix(matrixScale, 0d, 0d, matrixScale, 0d, 0d),
            tilingType: 2,
            "q 16 0 0 16 0 0 cm /" + imageName + " Do Q\n",
            [new PdfImageResource(imageName, image)]);
    }

    public static PdfTilingPattern OfficeBitmapStripeLines(
        PdfStripePatternKind kind,
        bool thin,
        byte foregroundRed,
        byte foregroundGreen,
        byte foregroundBlue,
        byte backgroundRed,
        byte backgroundGreen,
        byte backgroundBlue)
    {
        const int cellSize = 16;
        const double matrixScale = 0.375d;
        const string imageName = "ImPattern";
        int stripeWidth = thin ? 1 : 2;

        byte[] rgb = new byte[cellSize * cellSize * 3];
        for (int y = 0; y < cellSize; y++)
        {
            for (int x = 0; x < cellSize; x++)
            {
                int diagonalOffset = 2 * (y / 2);
                int position = kind switch
                {
                    PdfStripePatternKind.Horizontal => PositiveModulo(y, 8),
                    PdfStripePatternKind.Vertical => PositiveModulo(x, 8),
                    PdfStripePatternKind.UpDiagonal => PositiveModulo(x + diagonalOffset, 8),
                    PdfStripePatternKind.DownDiagonal => PositiveModulo(x - diagonalOffset, 8),
                    _ => 0
                };
                bool foreground = position >= 8 - stripeWidth;
                int index = (y * cellSize + x) * 3;
                rgb[index] = foreground ? foregroundRed : backgroundRed;
                rgb[index + 1] = foreground ? foregroundGreen : backgroundGreen;
                rgb[index + 2] = foreground ? foregroundBlue : backgroundBlue;
            }
        }

        var image = PdfImageXObject.RgbPng(cellSize, cellSize, rgb, alpha: null);
        return new PdfTilingPattern(
            cellSize,
            cellSize,
            cellSize,
            cellSize,
            new PdfPatternMatrix(matrixScale, 0d, 0d, matrixScale, 0d, 0d),
            tilingType: 2,
            "q 16 0 0 16 0 0 cm /" + imageName + " Do Q\n",
            [new PdfImageResource(imageName, image)]);
    }

    // Office pct70 percentage weave: two banked Office renders emit byte-identical 16x16
    // tiles (6pt at the 0.375 matrix scale, 192 foreground cells) with background cells where
    // ((x - 4 * band) mod 8) < 2 for 2-row band parity. Other densities keep the dot grid
    // (unobserved).
    public static PdfTilingPattern OfficeBitmapWeaveBlocks(
        byte foregroundRed,
        byte foregroundGreen,
        byte foregroundBlue,
        byte backgroundRed,
        byte backgroundGreen,
        byte backgroundBlue)
    {
        const int cellSize = 16;
        const double matrixScale = 0.375d;
        const string imageName = "ImPattern";

        byte[] rgb = new byte[cellSize * cellSize * 3];
        for (int y = 0; y < cellSize; y++)
        {
            for (int x = 0; x < cellSize; x++)
            {
                bool foreground = PositiveModulo(x - 4 * ((y / 2) % 2), 8) >= 2;
                int index = (y * cellSize + x) * 3;
                rgb[index] = foreground ? foregroundRed : backgroundRed;
                rgb[index + 1] = foreground ? foregroundGreen : backgroundGreen;
                rgb[index + 2] = foreground ? foregroundBlue : backgroundBlue;
            }
        }

        var image = PdfImageXObject.RgbPng(cellSize, cellSize, rgb, alpha: null);
        return new PdfTilingPattern(
            cellSize,
            cellSize,
            cellSize,
            cellSize,
            new PdfPatternMatrix(matrixScale, 0d, 0d, matrixScale, 0d, 0d),
            tilingType: 2,
            "q 16 0 0 16 0 0 cm /" + imageName + " Do Q\n",
            [new PdfImageResource(imageName, image)]);
    }

    // A preset pattern fill tile: the 8x8 bitmap (one byte per row, high bit leftmost)
    // doubled to the 16x16 image and 0.375 matrix scale Office emits, so one bitmap
    // pixel covers 0.75pt.
    public static PdfTilingPattern OfficeBitmapPreset(
        IReadOnlyList<byte> bitmapRows,
        byte foregroundRed,
        byte foregroundGreen,
        byte foregroundBlue,
        byte backgroundRed,
        byte backgroundGreen,
        byte backgroundBlue)
    {
        const int cellSize = 16;
        const double matrixScale = 0.375d;
        const string imageName = "ImPattern";

        byte[] rgb = new byte[cellSize * cellSize * 3];
        for (int y = 0; y < cellSize; y++)
        {
            for (int x = 0; x < cellSize; x++)
            {
                bool foreground = (bitmapRows[y / 2] & (0x80 >> (x / 2))) != 0;
                int index = (y * cellSize + x) * 3;
                rgb[index] = foreground ? foregroundRed : backgroundRed;
                rgb[index + 1] = foreground ? foregroundGreen : backgroundGreen;
                rgb[index + 2] = foreground ? foregroundBlue : backgroundBlue;
            }
        }

        var image = PdfImageXObject.RgbPng(cellSize, cellSize, rgb, alpha: null);
        return new PdfTilingPattern(
            cellSize,
            cellSize,
            cellSize,
            cellSize,
            new PdfPatternMatrix(matrixScale, 0d, 0d, matrixScale, 0d, 0d),
            tilingType: 2,
            "q 16 0 0 16 0 0 cm /" + imageName + " Do Q\n",
            [new PdfImageResource(imageName, image)]);
    }

    private string BuildResourceKey()
    {
        string imageKeys = string.Join("|", Images.Select(image => $"{image.ResourceName}:{image.Image.ResourceKey}"));
        return string.Create(CultureInfo.InvariantCulture, $"tiling:{Width:0.###}:{Height:0.###}:{XStep:0.###}:{YStep:0.###}:{Matrix}:{TilingType}:{Content}:{imageKeys}");
    }

    private static int PositiveModulo(int value, int modulus)
    {
        int result = value % modulus;
        return result < 0 ? result + modulus : result;
    }
}

internal readonly record struct PdfPatternMatrix(double A, double B, double C, double D, double E, double F)
{
    public override string ToString()
    {
        return string.Create(CultureInfo.InvariantCulture, $"{A:0.###}:{B:0.###}:{C:0.###}:{D:0.###}:{E:0.###}:{F:0.###}");
    }
}

internal enum PdfStripePatternKind
{
    Horizontal,
    Vertical,
    UpDiagonal,
    DownDiagonal
}
