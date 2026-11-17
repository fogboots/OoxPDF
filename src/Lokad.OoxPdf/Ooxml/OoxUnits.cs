using System.Globalization;

namespace Lokad.OoxPdf.Ooxml;

internal static class OoxUnits
{
    public const double EmusPerInch = 914400d;
    public const double TwipsPerPoint = 20d;
    public const double PointsPerInch = 72d;

    public static double EmuToPoints(long emu)
    {
        return emu / EmusPerInch * PointsPerInch;
    }

    public static double TwipsToPoints(long twips)
    {
        return twips / TwipsPerPoint;
    }

    /// <summary>
    /// Reads ST_TwipsMeasure / ST_SignedTwipsMeasure: an integer, or (ISO 29500 second edition
    /// on) a number with a universal-measure unit (mm, cm, in, pt, pc, pi). Producers other than
    /// Word also write plain decimals. Returns false for anything else (empty, "auto", text):
    /// callers keep their default, as Word's repair and LibreOffice do, instead of failing the
    /// whole document over one spacing value.
    /// </summary>
    public static bool TryParseTwipsMeasure(string? text, out double twips)
    {
        twips = 0d;
        if (text is null)
        {
            return false;
        }

        if (long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out long whole))
        {
            twips = whole;
            return true;
        }

        string value = text.Trim();
        double twipsPerUnit = 1d;
        if (value.Length > 2 && char.IsAsciiLetter(value[^1]) && char.IsAsciiLetter(value[^2]))
        {
            twipsPerUnit = value[^2..].ToLowerInvariant() switch
            {
                "pt" => 20d,
                "in" => 1440d,
                "cm" => 1440d / 2.54d,
                "mm" => 144d / 2.54d,
                "pc" or "pi" => 240d,
                _ => double.NaN
            };
            value = value[..^2];
        }

        if (double.IsNaN(twipsPerUnit) ||
            !TryParseDecimal(value, out double number) ||
            !double.IsFinite(number * twipsPerUnit))
        {
            return false;
        }

        twips = number * twipsPerUnit;
        return true;
    }

    /// <summary>The measure in points, or null when it cannot be read (see <see cref="TryParseTwipsMeasure"/>).</summary>
    public static double? TwipsMeasureToPoints(string? text)
    {
        return TryParseTwipsMeasure(text, out double twips) ? twips / TwipsPerPoint : null;
    }

    /// <summary>
    /// Reads an integer-valued attribute that some producers write as a decimal ("22.0"),
    /// within the int range the integer readers accepted. Returns false for anything else.
    /// </summary>
    public static bool TryParseDecimal(string? text, out double value)
    {
        value = 0d;
        return text is not null &&
            double.TryParse(text, NumberStyles.AllowLeadingWhite | NumberStyles.AllowTrailingWhite | NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out value) &&
            Math.Abs(value) <= int.MaxValue;
    }

    /// <summary>A whole number written as an integer or a decimal, truncated; false otherwise.</summary>
    public static bool TryParseWholeNumber(string? text, out int value)
    {
        value = 0;
        if (!TryParseDecimal(text, out double number))
        {
            return false;
        }

        value = (int)Math.Truncate(number);
        return true;
    }


}
