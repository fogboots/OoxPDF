using System.Globalization;
using System.Text;
using System.Xml.Linq;
using Lokad.OoxPdf.Diagnostics;
using Lokad.OoxPdf.Fonts;
using Lokad.OoxPdf.Ooxml;
using static Lokad.OoxPdf.Ooxml.OoxNamespaces;
using Lokad.OoxPdf.Pdf;

namespace Lokad.OoxPdf.Pptx;

internal sealed partial class PptxRenderer
{
    private static string FormatChartAxisLabel(double value, XElement? axis)
    {
        ChartNumberFormat numberFormat = axis is null ? default : ToChartNumberFormat(PptxSceneBuilder.ReadChartNumberFormat(axis));
        if (IsRenderableChartNumberFormat(numberFormat))
        {
            return FormatChartNumber(value, numberFormat.FormatCode, ResolveAxisDate1904(axis));
        }

        // Unrenderable codes fall back to General semantics: up to 10 significant digits
        // (Office renders 6.438276615812609 as 6.438276616), not two-decimal rounding.
        double rounded = Math.Round(value);
        return Math.Abs(value - rounded) < PptxChartMetricRules.AxisValueEpsilon
            ? rounded.ToString("0", CultureInfo.InvariantCulture)
            : value.ToString("G10", CultureInfo.InvariantCulture);
    }

    private static string FormatSceneOrXmlChartAxisLabel(double value, PptxSceneChartAxis? sceneAxis, XElement? axis, string? defaultNumberFormat)
    {
        ChartNumberFormat numberFormat = ReadSceneOrXmlChartAxisNumberFormat(sceneAxis, axis);
        if (IsRenderableChartNumberFormat(numberFormat))
        {
            return FormatChartNumber(value, numberFormat.FormatCode, ResolveAxisDate1904(axis));
        }

        return !string.IsNullOrWhiteSpace(defaultNumberFormat)
            ? FormatChartNumber(value, defaultNumberFormat)
            : FormatChartAxisLabel(value, null);
    }

    private static ChartNumberFormat ReadSceneOrXmlChartAxisNumberFormat(PptxSceneChartAxis? sceneAxis, XElement? axis)
    {
        if (sceneAxis is not null)
        {
            return ToChartNumberFormat(sceneAxis.NumberFormatInfo);
        }

        return axis is null ? default : ToChartNumberFormat(PptxSceneBuilder.ReadChartNumberFormat(axis));
    }

    private static bool IsRenderableChartNumberFormat(ChartNumberFormat numberFormat)
    {
        return numberFormat.IsDefined &&
            IsRenderableChartFormatCode(numberFormat.FormatCode);
    }

    private static bool IsRenderableChartFormatCode([System.Diagnostics.CodeAnalysis.NotNullWhen(true)] string? formatCode)
    {
        return !string.IsNullOrWhiteSpace(formatCode) &&
            !string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase);
    }

    private static string FormatChartNumber(double value, string formatCode, bool date1904 = false)
    {
        // Sections select by condition (first match or unconditional else), else by sign; a fourth text section never applies to numbers (S04).
        string[] sections = SplitChartNumberFormatSections(formatCode);
        int sectionIndex = SelectSectionIndex(sections, value);
        if (sectionIndex < 0)
        {
            return string.Empty;
        }

        string section = StripSectionCondition(sections[sectionIndex]);
        string localeCurrency = section.IndexOf("[", StringComparison.Ordinal) >= 0 ? ParseChartNumberLocaleCurrency(section) : string.Empty;
        // Calibrated [$LCID] brackets render locale decimal/group separators (S04); every other
        // bracket keeps invariant formatting with its diagnostic.
        CultureInfo numberCulture = CultureInfo.InvariantCulture;
        if (section.IndexOf("[$", StringComparison.Ordinal) >= 0
            && TryResolveChartNumberLocaleSeparators(section, out string localeDecimalSeparator, out string localeGroupSeparator)
            && (localeDecimalSeparator != "." || localeGroupSeparator != ","))
        {
            numberCulture = (CultureInfo)CultureInfo.InvariantCulture.Clone();
            numberCulture.NumberFormat.NumberDecimalSeparator = localeDecimalSeparator;
            numberCulture.NumberFormat.NumberGroupSeparator = localeGroupSeparator;
        }
        if (TryFormatDateNumber(section, value, date1904, out string? dateText))
        {
            return localeCurrency.Length == 0 ? dateText : localeCurrency + dateText;
        }

        ExtractChartNumberSectionLiterals(section, out string literalPrefix, out string literalSuffix, out bool hasDigitPlaceholders);
        if (!hasDigitPlaceholders && (literalPrefix.Length > 0 || literalSuffix.Length > 0))
        {
            return literalPrefix + literalSuffix;
        }

        string view = ChartNumberFormatSyntaxView(section);
        bool percent = view.Contains('%', StringComparison.Ordinal);
        double displayValue = percent ? value * 100d : value;
        int scalingCommas = 0;
        for (int i = view.Length - 1; i >= 0 && view[i] == ','; i--)
        {
            scalingCommas++;
        }

        for (int i = 0; i < scalingCommas; i++)
        {
            displayValue /= 1000d;
        }

        int decimals = 0;
        int decimalPoint = view.IndexOf('.', StringComparison.Ordinal);
        if (decimalPoint >= 0)
        {
            for (int i = decimalPoint + 1; i < view.Length && view[i] is '0' or '#'; i++)
            {
                decimals++;
            }
        }

        bool thousands = false;
        for (int i = 0; i < view.Length - scalingCommas; i++)
        {
            if (view[i] == ',')
            {
                thousands = true;
                break;
            }
        }

        // Explicit negative/zero sections carry their own sign; single-section formats keep signed .NET formatting.
        double emitValue = sectionIndex > 0 ? Math.Abs(displayValue) : displayValue;
        string text;
        if (TryFormatScientific(view, emitValue, numberCulture, out string? scientific))
        {
            text = scientific;
        }
        else if (TryFormatFraction(view, emitValue, numberCulture, out string? fraction))
        {
            text = fraction;
        }
        else
        {
            string numberFormat = (thousands ? "#,##0" : "0") +
                (decimals > 0 ? "." + new string('0', decimals) : string.Empty);
            text = emitValue.ToString(numberFormat, numberCulture);
        }

        if (view.Contains('$', StringComparison.Ordinal) || localeCurrency.Length > 0 || view.Contains(LiteralDollarPlaceholder, StringComparison.Ordinal))
        {
            text = (localeCurrency.Length > 0 ? localeCurrency : "$") + text;

        }
        if (sectionIndex == 1)
        {
            if (view.Contains('(', StringComparison.Ordinal) && view.Contains(')', StringComparison.Ordinal))
            {
                text = "(" + text + ")";
            }
            else if (ChartNumberFormatSectionStartsWithMinus(section))
            {
                text = "-" + text;
            }
        }

        if (percent || view.Contains(LiteralPercentPlaceholder, StringComparison.Ordinal))
        {
            text += "%";
        }

        return literalPrefix.Length == 0 && literalSuffix.Length == 0 ? text : literalPrefix + text + literalSuffix;
    }

    // Scientific notation renders mantissa and base-10 exponent from the digit
    // placeholders around E[+|-]; escaped/quoted characters are already masked
    // in the syntax view (S04).
    private static bool TryFormatScientific(string view, double value, CultureInfo numberCulture, out string text)
    {
        text = string.Empty;
        int eIndex = -1;
        for (int i = 0; i < view.Length; i++)
        {
            if (view[i] is not ('E' or 'e'))
            {
                continue;
            }

            if (HasMantissaPlaceholder(view, i) && ExponentWidth(view, i) > 0)
            {
                eIndex = i;
                break;
            }
        }

        if (eIndex < 0)
        {
            return false;
        }

        int decimals = 0;
        int decimalPoint = view.IndexOf('.', StringComparison.Ordinal);
        if (decimalPoint >= 0 && decimalPoint < eIndex)
        {
            for (int i = decimalPoint + 1; i < eIndex && view[i] is '0' or '#'; i++)
            {
                decimals++;
            }
        }

        int intDigits = 0;
        for (int i = 0; i < eIndex && view[i] != '.'; i++)
        {
            if (view[i] is '0' or '#' or '?')
            {
                intDigits++;
            }
        }

        intDigits = Math.Max(1, intDigits);
        bool alwaysSign = eIndex + 1 < view.Length && view[eIndex + 1] == '+';
        int expWidth = ExponentWidth(view, eIndex);
        double magnitude = Math.Abs(value);
        int exponent;
        double mantissa;
        if (magnitude == 0d)
        {
            exponent = 0;
            mantissa = 0d;
        }
        else
        {
            exponent = (int)Math.Floor(Math.Log10(magnitude)) - intDigits + 1;
            mantissa = magnitude / Math.Pow(10d, exponent);
            double scale = Math.Pow(10d, decimals);
            mantissa = Math.Round(mantissa * scale, MidpointRounding.AwayFromZero) / scale;
            if (mantissa >= Math.Pow(10d, intDigits))
            {
                mantissa /= 10d;
                exponent++;
            }
        }

        string mantissaPattern = new string('0', intDigits) + (decimals > 0 ? "." + new string('0', decimals) : string.Empty);
        string mantissaText = (value < 0d ? -mantissa : mantissa).ToString(mantissaPattern, numberCulture);
        string sign = exponent < 0 ? "-" : alwaysSign ? "+" : string.Empty;
        text = mantissaText + view[eIndex] + sign + Math.Abs(exponent).ToString(CultureInfo.InvariantCulture).PadLeft(expWidth, '0');
        return true;
    }

    private static bool HasMantissaPlaceholder(string view, int eIndex)
    {
        for (int i = 0; i < eIndex; i++)
        {
            if (view[i] is '0' or '#' or '?' or '.')
            {
                return true;
            }
        }

        return false;
    }

    private static int ExponentWidth(string view, int eIndex)
    {
        int i = eIndex + 1;
        if (i < view.Length && view[i] is '+' or '-')
        {
            i++;
        }

        int width = 0;
        while (i < view.Length && view[i] is '0' or '#')
        {
            width++;
            i++;
        }

        return width;
    }

    // Fractions render whole and numerator/denominator from ? precision; the
    // denominator search caps at 9999 to bound per-label work (S04).
    private static bool TryFormatFraction(string view, double emitValue, CultureInfo numberCulture, out string text)
    {
        text = string.Empty;
        int slash = view.IndexOf('/', StringComparison.Ordinal);
        if (slash < 0)
        {
            return false;
        }

        int denDigits = 0;
        for (int i = slash + 1; i < view.Length && (view[i] == '?' || view[i] == ' '); i++)
        {
            if (view[i] == '?')
            {
                denDigits++;
            }
        }

        if (denDigits == 0)
        {
            return false;
        }

        int numStart = slash;
        while (numStart > 0 && view[numStart - 1] is '0' or '#' or '?')
        {
            numStart--;
        }

        int intEnd = numStart;
        while (intEnd > 0 && view[intEnd - 1] == ' ')
        {
            intEnd--;
        }

        int intStart = intEnd;
        while (intStart > 0 && view[intStart - 1] is '0' or '#' or '?')
        {
            intStart--;
        }

        bool showWhole = false;
        for (int i = intStart; i < intEnd; i++)
        {
            if (view[i] == '0')
            {
                showWhole = true;
                break;
            }
        }

        int maxDenominator = 1;
        for (int i = 0; i < denDigits && maxDenominator <= 999; i++)
        {
            maxDenominator *= 10;
        }

        maxDenominator -= 1;
        double magnitude = Math.Abs(emitValue);
        int whole = (int)Math.Floor(magnitude);
        double remainder = magnitude - whole;
        int bestNum = 0;
        int bestDen = 1;
        double bestError = remainder;
        for (int den = 1; den <= maxDenominator; den++)
        {
            int num = (int)Math.Round(remainder * den, MidpointRounding.AwayFromZero);
            if (num < 0 || num > den)
            {
                continue;
            }

            double error = Math.Abs(remainder - (double)num / den);
            if (error < bestError - 1e-12)
            {
                bestError = error;
                bestNum = num;
                bestDen = den;
            }
        }

        if (bestNum >= bestDen)
        {
            whole++;
            bestNum = 0;
        }

        string sign = emitValue < 0d ? "-" : string.Empty;
        if (bestNum == 0)
        {
            text = sign + whole.ToString(CultureInfo.InvariantCulture);
            return true;
        }

        text = sign + (showWhole || whole != 0 ? whole.ToString(numberCulture) + " " : string.Empty)
            + bestNum.ToString(numberCulture) + "/" + bestDen.ToString(numberCulture);
        return true;
    }

    // Chart-space date system for axis labels, resolved from the axis XML
    // ancestry so no render signature changes (S04). Defaults to 1900.
    private static bool ResolveAxisDate1904(XElement? axis)
    {
        return ResolveChartSpaceDate1904(axis);
    }

    private static bool ResolveChartSpaceDate1904(XElement? element)
    {
        if (element is null)
        {
            return false;
        }

        XElement? chartSpace = element.Name == ChartNamespace + "chartSpace"
            ? element
            : element.Ancestors(ChartNamespace + "chartSpace").FirstOrDefault();
        XElement? dateElement = chartSpace?.Element(ChartNamespace + "date1904");
        return dateElement is not null && PptxSceneBuilder.IsOoxmlBooleanElementEnabled(dateElement);
    }

    private enum DateFormatTokenKind
    {
        Year,
        Month,
        Minutes,
        Day,
        Hour,
        Second,
        AmPm,
        ElapsedHour,
        ElapsedMinute,
        ElapsedSecond,
        Literal
    }

    private readonly record struct DateFormatToken(DateFormatTokenKind Kind, string Text, int FractionDigits = 0);

    private static readonly string[] EnglishMonthNames =
    [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ];

    private static readonly string[] EnglishDayNames =
    [
        "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"
    ];

    // Serial date/datetime rendering for axis labels (S04). Only the 1900 and
    // 1904 systems are supported; English month/day names (locale stays open).
    // Unrepresentable serials fall back to numeric rendering. Serial 0 renders as the
    // phantom January 0, and serial 60 as March 1 (the phantom Feb 29 is unrepresentable).
    private static bool TryFormatDateNumber(string section, double value, bool date1904, out string text)
    {
        text = string.Empty;
        if (double.IsNaN(value) || double.IsInfinity(value))
        {
            return false;
        }

        List<DateFormatToken>? tokens = TokenizeDateFormat(section);
        if (tokens is null)
        {
            return false;
        }

        double wholeDays = Math.Floor(value);
        DateOnly origin = date1904 ? new DateOnly(1904, 1, 1) : new DateOnly(1899, 12, 30);
        if (wholeDays < 0d || wholeDays > DateOnly.MaxValue.DayNumber - origin.DayNumber)
        {
            return false;
        }

        int days = (int)wholeDays;
        double fraction = value - days;
        int year;
        int month;
        int day;
        DayOfWeek weekday;
        if (date1904)
        {
            DateOnly date = new DateOnly(1904, 1, 1).AddDays(days);
            year = date.Year;
            month = date.Month;
            day = date.Day;
            weekday = date.DayOfWeek;
        }
        else
        {
            if (days == 0)
            {
                // Phantom January 0 displays with day 0 (S04).
                year = 1900;
                month = 1;
                day = 0;
                weekday = DayOfWeek.Saturday;
            }
            else
            {
                DateOnly date = days < 61
                    ? new DateOnly(1899, 12, 31).AddDays(days)
                    : new DateOnly(1899, 12, 30).AddDays(days);
                year = date.Year;
                month = date.Month;
                day = date.Day;
                // Lotus leap-year bug: Excel weekday names shift by one below March 1900.
                weekday = days < 61 ? (DayOfWeek)(((int)date.DayOfWeek + 6) % 7) : date.DayOfWeek;
            }
        }

        long totalMs = (long)(fraction * 86400000d);
        int hour24 = (int)(totalMs / 3600000L);
        int minute = (int)(totalMs / 60000L % 60L);
        int second = (int)(totalMs / 1000L % 60L);
        int milli = (int)(totalMs % 1000L);
        bool hasAmPm = tokens.Exists(token => token.Kind == DateFormatTokenKind.AmPm);
        var builder = new StringBuilder();
        foreach (DateFormatToken token in tokens)
        {
            switch (token.Kind)
            {
                case DateFormatTokenKind.Year:
                    builder.Append(token.Text.Length >= 3 ? year.ToString(CultureInfo.InvariantCulture) : (year % 100).ToString("00", CultureInfo.InvariantCulture));
                    break;
                case DateFormatTokenKind.Month:
                    AppendMonth(builder, month, token.Text.Length);
                    break;
                case DateFormatTokenKind.Minutes:
                    AppendPadded(builder, minute, token.Text.Length);
                    break;
                case DateFormatTokenKind.Day:
                    AppendDay(builder, weekday, day, token.Text.Length);
                    break;
                case DateFormatTokenKind.Hour:
                    AppendPadded(builder, hasAmPm ? ToTwelveHour(hour24) : hour24, token.Text.Length);
                    break;
                case DateFormatTokenKind.Second:
                    AppendPadded(builder, second, token.Text.Length);
                    if (token.FractionDigits > 0)
                    {
                        int divisor = 1;
                        for (int i = 0; i < 3 - Math.Min(3, token.FractionDigits); i++)
                        {
                            divisor *= 10;
                        }

                        int digits = token.FractionDigits >= 3 ? milli : milli / divisor;
                        builder.Append('.');
                        builder.Append(digits.ToString(CultureInfo.InvariantCulture).PadLeft(token.FractionDigits, '0'));
                    }

                    break;
                case DateFormatTokenKind.AmPm:
                    builder.Append(FormatMeridiem(hour24 < 12, token.Text));
                    break;
                case DateFormatTokenKind.ElapsedHour:
                    builder.Append(((long)days * 24L + hour24).ToString(CultureInfo.InvariantCulture));
                    break;
                case DateFormatTokenKind.ElapsedMinute:
                    builder.Append(((long)days * 1440L + hour24 * 60L + minute).ToString(CultureInfo.InvariantCulture));
                    break;
                case DateFormatTokenKind.ElapsedSecond:
                    builder.Append((days * 86400L + totalMs / 1000L).ToString(CultureInfo.InvariantCulture));
                    break;
                default:
                    builder.Append(token.Text);
                    break;
            }
        }

        text = builder.ToString();
        return true;

        static void AppendMonth(StringBuilder builder, int month, int length)
        {
            if (length == 3)
            {
                builder.Append(EnglishMonthNames[month - 1].Substring(0, 3));
            }
            else if (length == 4)
            {
                builder.Append(EnglishMonthNames[month - 1]);
            }
            else if (length >= 5)
            {
                builder.Append(EnglishMonthNames[month - 1][0]);
            }
            else
            {
                AppendPadded(builder, month, length);
            }
        }

        static void AppendDay(StringBuilder builder, DayOfWeek weekday, int day, int length)
        {
            if (length == 3)
            {
                builder.Append(EnglishDayNames[(int)weekday].Substring(0, 3));
            }
            else if (length >= 4)
            {
                builder.Append(EnglishDayNames[(int)weekday]);
            }
            else
            {
                AppendPadded(builder, day, length);
            }
        }

        static void AppendPadded(StringBuilder builder, int part, int length)
        {
            builder.Append(part.ToString(CultureInfo.InvariantCulture).PadLeft(Math.Min(2, length), '0'));
        }

        static int ToTwelveHour(int hour24)
        {
            int hour = hour24 % 12;
            return hour == 0 ? 12 : hour;
        }

        static string FormatMeridiem(bool morning, string pattern)
        {
            bool upper = true;
            foreach (char c in pattern)
            {
                if (char.IsLower(c))
                {
                    upper = false;
                    break;
                }

                if (char.IsUpper(c))
                {
                    upper = true;
                    break;
                }
            }

            string meridiem = pattern.Length == 1
                ? (morning ? "A" : "P")
                : (morning ? "AM" : "PM");
            return upper ? meridiem : meridiem.ToLowerInvariant();
        }
    }

    // Splits a format section into date tokens and literals; returns null when
    // no date token exists. Bracketed colors/conditions/locales are skipped
    // (except elapsed h/m/s); quotes, escapes, and cushion/fill runs are kept
    // or skipped exactly like the syntax view (S04).
    private static List<DateFormatToken>? TokenizeDateFormat(string section)
    {
        var tokens = new List<DateFormatToken>();
        bool hasDateToken = false;
        int i = 0;
        while (i < section.Length)
        {
            char c = section[i];
            if (c == '\\' && i + 1 < section.Length)
            {
                tokens.Add(new DateFormatToken(DateFormatTokenKind.Literal, section.Substring(i + 1, 1)));
                i += 2;
                continue;
            }

            if (c == '"')
            {
                int close = section.IndexOf('"', i + 1);
                if (close < 0)
                {
                    tokens.Add(new DateFormatToken(DateFormatTokenKind.Literal, section.Substring(i + 1).Replace("\"\"", "\"", StringComparison.Ordinal)));
                    break;
                }

                tokens.Add(new DateFormatToken(DateFormatTokenKind.Literal, section.Substring(i + 1, close - i - 1).Replace("\"\"", "\"", StringComparison.Ordinal)));
                i = close + 1;
                continue;
            }

            if (c == '[')
            {
                int close = section.IndexOf(']', i);
                if (close < 0)
                {
                    tokens.Add(new DateFormatToken(DateFormatTokenKind.Literal, c.ToString()));
                    i++;
                    continue;
                }

                string inner = section.Substring(i + 1, close - i - 1);
                if (inner.Length == 1 && (inner[0] is 'h' or 'H'))
                {
                    tokens.Add(new DateFormatToken(DateFormatTokenKind.ElapsedHour, section.Substring(i, close - i + 1)));
                    hasDateToken = true;
                }
                else if (inner.Length >= 1 && inner.Length <= 2 && inner.All(ch => ch is 'm' or 'M'))
                {
                    tokens.Add(new DateFormatToken(DateFormatTokenKind.ElapsedMinute, section.Substring(i, close - i + 1)));
                    hasDateToken = true;
                }
                else if ((inner.Length == 1 && (inner[0] is 's' or 'S')) || (inner.Length == 2 && (inner[0] is 's' or 'S') && (inner[1] is 's' or 'S')))
                {
                    tokens.Add(new DateFormatToken(DateFormatTokenKind.ElapsedSecond, section.Substring(i, close - i + 1)));
                    hasDateToken = true;
                }

                i = close + 1;
                continue;
            }

            if ((c == '_' || c == '*') && i + 1 < section.Length)
            {
                i += 2;
                continue;
            }

            if (c == '_' || c == '*')
            {
                i++;
                continue;
            }

            if (IsAsciiLetter(c))
            {
                int runEnd = i;
                while (runEnd < section.Length && IsAsciiLetter(section[runEnd]))
                {
                    runEnd++;
                }

                string run = section.Substring(i, runEnd - i);
                if (TryMatchDateRun(run, out DateFormatTokenKind kind, out int used))
                {
                    string head = section.Substring(i, used);
                    int fractionDigits = 0;
                    if (kind == DateFormatTokenKind.Second)
                    {
                        int j = i + used;
                        if (j < section.Length && section[j] == '.')
                        {
                            int k = j + 1;
                            while (k < section.Length && section[k] == '0')
                            {
                                k++;
                            }

                            if (k > j + 1)
                            {
                                fractionDigits = k - j - 1;
                                used = k - i;
                                head = section.Substring(i, used);
                            }
                        }
                    }

                    tokens.Add(new DateFormatToken(kind, head, fractionDigits));
                    hasDateToken = true;
                    i += used;
                    continue;
                }

                tokens.Add(new DateFormatToken(DateFormatTokenKind.Literal, run));
                i = runEnd;
                continue;
            }

            tokens.Add(new DateFormatToken(DateFormatTokenKind.Literal, c.ToString()));
            i++;
        }

        if (!hasDateToken)
        {
            return null;
        }

        for (int k = 0; k < tokens.Count; k++)
        {
            if (tokens[k].Kind != DateFormatTokenKind.Month)
            {
                continue;
            }

            if ((PreviousDateToken(tokens, k) is { Kind: DateFormatTokenKind.Hour or DateFormatTokenKind.ElapsedHour }) ||
                (NextDateToken(tokens, k) is { Kind: DateFormatTokenKind.Second or DateFormatTokenKind.ElapsedSecond }))
            {
                tokens[k] = tokens[k] with { Kind = DateFormatTokenKind.Minutes };
            }
        }

        return tokens;

        static DateFormatToken? PreviousDateToken(List<DateFormatToken> tokens, int index)
        {
            for (int k = index - 1; k >= 0; k--)
            {
                if (tokens[k].Kind != DateFormatTokenKind.Literal)
                {
                    return tokens[k];
                }
            }

            return null;
        }

        static DateFormatToken? NextDateToken(List<DateFormatToken> tokens, int index)
        {
            for (int k = index + 1; k < tokens.Count; k++)
            {
                if (tokens[k].Kind != DateFormatTokenKind.Literal)
                {
                    return tokens[k];
                }
            }

            return null;
        }
    }

    private static bool IsAsciiLetter(char c)
    {
        return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
    }

    // A letter run is a date run only when it fully matches date-token grammar;
    // ordinary words (days, items, Sep) stay literal (S04).
    private static bool TryMatchDateRun(string run, out DateFormatTokenKind kind, out int used)
    {
        kind = DateFormatTokenKind.Literal;
        used = 0;
        if (run.Length >= 2 && (run.StartsWith("am", StringComparison.OrdinalIgnoreCase) || run.StartsWith("pm", StringComparison.OrdinalIgnoreCase)))
        {
            kind = DateFormatTokenKind.AmPm;
            used = 2;
            return true;
        }

        char first = char.ToLowerInvariant(run[0]);
        if (first is 'y' or 'm' or 'd' or 'h' or 's' or 'a' or 'p')
        {
            int length = 0;
            while (length < run.Length && char.ToLowerInvariant(run[length]) == first)
            {
                length++;
            }

            if (first is 'a' or 'p')
            {
                if (length > 1)
                {
                    return false;
                }

                kind = DateFormatTokenKind.AmPm;
                used = 1;
                return true;
            }

            kind = first switch
            {
                'y' => DateFormatTokenKind.Year,
                'm' => DateFormatTokenKind.Month,
                'd' => DateFormatTokenKind.Day,
                'h' => DateFormatTokenKind.Hour,
                _ => DateFormatTokenKind.Second,
            };
            used = length;
            return true;
        }

        return false;
    }

    // Section selection: formats carrying [op value] conditions pick the first
    // matching section (an unconditional section is the else branch); with no
    // match nothing renders. Plain formats keep sign selection (S04).
    private static int SelectSectionIndex(string[] sections, double value)
    {
        bool anyCondition = false;
        foreach (string section in sections)
        {
            if (TryParseSectionCondition(section, out _, out _, out _))
            {
                anyCondition = true;
                break;
            }
        }

        if (!anyCondition)
        {
            return value < 0d && sections.Length > 1 ? 1 : value == 0d && sections.Length > 2 ? 2 : 0;
        }

        for (int i = 0; i < sections.Length; i++)
        {
            if (!TryParseSectionCondition(sections[i], out string op, out double threshold, out _))
            {
                return i;
            }

            if (EvaluateSectionCondition(op, threshold, value))
            {
                return i;
            }
        }

        return -1;
    }

    private static bool TryParseSectionCondition(string section, out string op, out double threshold, out int prefixLength)
    {
        op = string.Empty;
        threshold = 0d;
        prefixLength = 0;
        if (!section.StartsWith('['))
        {
            return false;
        }

        int close = section.IndexOf(']');
        if (close < 0)
        {
            return false;
        }

        string inner = section.Substring(1, close - 1).Trim();
        string[] operators = ["==", ">=", "<=", "<>", ">", "<", "="];
        foreach (string candidate in operators)
        {
            if (inner.StartsWith(candidate, StringComparison.Ordinal))
            {
                string number = inner.Substring(candidate.Length).Trim();
                if (double.TryParse(number, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double parsed))
                {
                    op = candidate.Length == 2 && candidate[1] == '=' && candidate[0] != '!' && candidate[0] != '<' && candidate[0] != '>' ? "=" : candidate;
                    threshold = parsed;
                    prefixLength = close + 1;
                    return true;
                }

                return false;
            }
        }

        return false;
    }

    private static bool EvaluateSectionCondition(string op, double threshold, double value)
    {
        return op switch
        {
            ">" => value > threshold,
            ">=" => value >= threshold,
            "<" => value < threshold,
            "<=" => value <= threshold,
            "=" or "==" => value == threshold,
            "<>" => value != threshold,
            _ => false,
        };
    }

    private static string StripSectionCondition(string section)
    {
        return TryParseSectionCondition(section, out _, out _, out int prefixLength)
            ? section.Substring(prefixLength)
            : section;
    }

    // Leading sign of a negative section, skipping quotes, escapes, fill/skip runs, and color/condition brackets (S04).
    private static bool ChartNumberFormatSectionStartsWithMinus(string section)
    {
        for (int i = 0; i < section.Length; i++)
        {
            char c = section[i];
            if (c is ' ' or '\t')
            {
                continue;
            }

            if (c == '"' || (c == '\\' && i + 1 < section.Length))
            {
                continue;
            }

            if ((c == '_' || c == '*') && i + 1 < section.Length)
            {
                i++;
                continue;
            }

            if (c == '[')
            {
                int close = section.IndexOf(']', i);
                if (close < 0)
                {
                    return false;
                }

                i = close;
                continue;
            }

            return c == '-';
        }

        return false;
    }
    private const char LiteralPercentPlaceholder = '\u0002';

    private const char LiteralDollarPlaceholder = '\u0003';

    private static string[] SplitChartNumberFormatSections(string formatCode)
    {
        var sections = new List<string>();
        var current = new StringBuilder();
        bool inQuotes = false;
        for (int i = 0; i < formatCode.Length; i++)
        {
            char c = formatCode[i];
            if (c == '\\' && i + 1 < formatCode.Length)
            {
                current.Append(c).Append(formatCode[++i]);
                continue;
            }

            if (c == '"')
            {
                inQuotes = !inQuotes;
                current.Append(c);
                continue;
            }

            if (c == ';' && !inQuotes)
            {
                sections.Add(current.ToString());
                current.Clear();
                continue;
            }

            current.Append(c);
        }

        sections.Add(current.ToString());
        return sections.ToArray();
    }

    // Literal runs in one numeric section (S04). Quoted ("..."/"") and escaped runs are
    // strong literals, except that quoted/escaped % and $ stay with the caller's
    // placeholder logic and escaped _/* runs are spacing like in the syntax view.
    // Unquoted letter/space/digit runs are weak literals that Excel also renders beside
    // digit placeholders; a weak run holding e/E/digits is dropped instead of partially
    // emitted, runs strictly inside the placeholder span (for example the fraction gap
    // in "# ?/?") are dropped, and bracket spans, cushion/fill runs, and other unquoted
    // syntax (including positional $/%) stay out of scope.
    private static void ExtractChartNumberSectionLiterals(string section, out string prefix, out string suffix, out bool hasDigitPlaceholders)
    {
        prefix = string.Empty;
        suffix = string.Empty;
        hasDigitPlaceholders = false;
        if (!ChartNumberSectionMayContainLiterals(section))
        {
            foreach (char scan in section)
            {
                if (scan is '0' or '#' or '?')
                {
                    hasDigitPlaceholders = true;
                    break;
                }
            }

            return;
        }

        StringBuilder? leading = null;
        StringBuilder? trailing = null;
        List<(int Start, int End, int StartPlaceholders)>? pendingWeak = null;
        int placeholders = 0;
        bool inQuotes = false;
        for (int i = 0; i < section.Length; i++)
        {
            char c = section[i];
            if (!inQuotes && c == '\\' && i + 1 < section.Length)
            {
                char escaped = section[++i];
                if (escaped is '_' or '*')
                {
                    continue;
                }

                if (escaped is '%' or '$')
                {
                    continue;
                }

                if (placeholders == 0)
                {
                    leading ??= new StringBuilder();
                    leading.Append(escaped);
                }
                else
                {
                    trailing ??= new StringBuilder();
                    trailing.Append(escaped);
                }

                continue;
            }

            if (c == '"')
            {
                if (inQuotes && i + 1 < section.Length && section[i + 1] == '"')
                {
                    if (placeholders == 0)
                    {
                        leading ??= new StringBuilder();
                        leading.Append('"');
                    }
                    else
                    {
                        trailing ??= new StringBuilder();
                        trailing.Append('"');
                    }

                    i++;
                    continue;
                }

                inQuotes = !inQuotes;
                continue;
            }

            if (inQuotes)
            {
                if (c is '%' or '$')
                {
                    continue;
                }

                if (placeholders == 0)
                {
                    leading ??= new StringBuilder();
                    leading.Append(c);
                }
                else
                {
                    trailing ??= new StringBuilder();
                    trailing.Append(c);
                }

                continue;
            }

            if (c == '[')
            {
                int close = section.IndexOf(']', i);
                if (close >= 0)
                {
                    i = close;
                }

                continue;
            }

            if ((c == '_' || c == '*') && i + 1 < section.Length)
            {
                i++;
                continue;
            }

            if (c is '0' or '#' or '?')
            {
                placeholders++;
                continue;
            }

            if (IsAsciiLetter(c) || c == ' ' || (c >= '1' && c <= '9'))
            {
                int runStart = i;
                bool poisoned = false;
                while (i < section.Length && (IsAsciiLetter(section[i]) || section[i] == ' ' || (section[i] >= '1' && section[i] <= '9')))
                {
                    if (section[i] is 'e' or 'E' || (section[i] >= '1' && section[i] <= '9'))
                    {
                        poisoned = true;
                    }

                    i++;
                }

                int runEnd = i - 1;
                i--;
                if (!poisoned)
                {
                    pendingWeak ??= new();
                    pendingWeak.Add((runStart, runEnd, placeholders));
                }

                continue;
            }
        }

        hasDigitPlaceholders = placeholders > 0;
        if (pendingWeak is not null && hasDigitPlaceholders)
        {
            foreach ((int start, int end, int startPlaceholders) in pendingWeak)
            {
                if (startPlaceholders == 0)
                {
                    leading ??= new StringBuilder();
                    leading.Append(section, start, end - start + 1);
                }
                else if (startPlaceholders == placeholders)
                {
                    trailing ??= new StringBuilder();
                    trailing.Append(section, start, end - start + 1);
                }
            }
        }

        if (leading is not null)
        {
            prefix = leading.ToString();
        }

        if (trailing is not null)
        {
            suffix = trailing.ToString();
        }
    }

    // True when the section can hold literal runs worth a full walk: quotes, escapes,
    // brackets, cushion/fill runs, or unquoted letters/spaces (S04).
    private static bool ChartNumberSectionMayContainLiterals(string section)
    {
        foreach (char c in section)
        {
            if ((c is '"' or '\\' or '[' or '_' or '*' or ' ') || IsAsciiLetter(c))
            {
                return true;
            }
        }

        return false;
    }

    // Unsupported-construct inventory for one chart number-format code (S04). The IDs
    // are categorical so document content never leaks into diagnostics; emission
    // through the diagnostic sink follows separately.
    internal static IReadOnlyList<string> GetUnsupportedChartNumberFormatConstructs(string formatCode)
    {
        if (string.IsNullOrWhiteSpace(formatCode) ||
            string.Equals(formatCode, "General", StringComparison.OrdinalIgnoreCase))
        {
            return [];
        }

        List<string>? unsupported = null;
        foreach (string section in SplitChartNumberFormatSections(formatCode))
        {
            CollectUnsupportedChartNumberSectionConstructs(section, ref unsupported);
        }

        return unsupported is null ? [] : unsupported;
    }

    private static void CollectUnsupportedChartNumberSectionConstructs(string section, ref List<string>? unsupported)
    {
        bool inQuotes = false;
        for (int i = 0; i < section.Length; i++)
        {
            char c = section[i];
            if (!inQuotes && c == '\\' && i + 1 < section.Length)
            {
                i++;
                continue;
            }

            if (c == '"')
            {
                inQuotes = !inQuotes;
                continue;
            }

            if (inQuotes)
            {
                continue;
            }

            if (c == '[')
            {
                int close = section.IndexOf(']', i);
                if (close < 0)
                {
                    continue;
                }

                ClassifyChartNumberFormatBracket(section, i, close, ref unsupported);
                i = close;
                continue;
            }

            if (c == '!')
            {
                AddUnsupportedChartConstruct(ref unsupported, "bang-escape");
                continue;
            }

            if (c == '@')
            {
                AddUnsupportedChartConstruct(ref unsupported, "text-placeholder");
            }
        }
    }

    private static void ClassifyChartNumberFormatBracket(string section, int bracketIndex, int bracketClose, ref List<string>? unsupported)
    {
        if (TryParseSectionCondition(section.Substring(bracketIndex), out _, out _, out _))
        {
            return;
        }

        string inner = section.Substring(bracketIndex + 1, bracketClose - bracketIndex - 1).Trim();
        if ((inner.Length == 1 && (inner[0] is 'h' or 'H')) ||
            (inner.Length >= 1 && inner.Length <= 2 && inner.All(static ch => ch is 'm' or 'M')) ||
            ((inner.Length == 1 && (inner[0] is 's' or 'S')) || (inner.Length == 2 && (inner[0] is 's' or 'S') && (inner[1] is 's' or 'S'))))
        {
            return;
        }

        if (inner.StartsWith('$'))
        {
            // Calibrated LCIDs render locale separators (no diagnostic); system and unlisted
            // LCIDs keep the "locale" fallback warning.
            if (TryParseChartNumberLocaleId(inner, out int localeId)
                && TryGetChartNumberLocaleSeparators(localeId, out _, out _))
            {
                return;
            }

            AddUnsupportedChartConstruct(ref unsupported, "locale");
            return;
        }

        if (IsChartNumberFormatColorName(inner))
        {
            AddUnsupportedChartConstruct(ref unsupported, "color");
            return;
        }

        if (inner.StartsWith("DBNum", StringComparison.OrdinalIgnoreCase) ||
            inner.StartsWith("NatNum", StringComparison.OrdinalIgnoreCase))
        {
            AddUnsupportedChartConstruct(ref unsupported, "native-digits");
            return;
        }

        AddUnsupportedChartConstruct(ref unsupported, "unknown-bracket");
    }

    private static bool IsChartNumberFormatColorName(string inner)
    {
        if (inner.Equals("Black", StringComparison.OrdinalIgnoreCase) ||
            inner.Equals("Blue", StringComparison.OrdinalIgnoreCase) ||
            inner.Equals("Cyan", StringComparison.OrdinalIgnoreCase) ||
            inner.Equals("Green", StringComparison.OrdinalIgnoreCase) ||
            inner.Equals("Magenta", StringComparison.OrdinalIgnoreCase) ||
            inner.Equals("Red", StringComparison.OrdinalIgnoreCase) ||
            inner.Equals("White", StringComparison.OrdinalIgnoreCase) ||
            inner.Equals("Yellow", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return inner.StartsWith("Color", StringComparison.OrdinalIgnoreCase) &&
            inner.Length > 5 &&
            inner.Substring(5).All(static ch => ch is >= '0' and <= '9');
    }

    private static void AddUnsupportedChartConstruct(ref List<string>? unsupported, string id)
    {
        unsupported ??= new List<string>();
        if (!unsupported.Contains(id))
        {
            unsupported.Add(id);
        }
    }

    // Diagnosed fallback for unsupported number-format syntax (S04): one diagnostic
    // per construct per chart, over scene formats with an XML fallback. Messages stay
    // categorical so document content never leaks into diagnostics.
    private static void EmitUnsupportedChartNumberFormatDiagnostics(
        XDocument chartXml,
        PptxSceneChart? sceneChart,
        Action<OoxPdfDiagnostic>? diagnosticSink,
        string? chartPartName,
        int slideIndex)
    {
        if (diagnosticSink is null)
        {
            return;
        }

        HashSet<string>? emitted = null;
        void collect(string? formatCode)
        {
            foreach (string construct in GetUnsupportedChartNumberFormatConstructs(formatCode ?? string.Empty))
            {
                if (emitted?.Contains(construct) == true)
                {
                    continue;
                }

                emitted ??= new HashSet<string>(StringComparer.Ordinal);
                emitted.Add(construct);
                EmitChartDiagnostic(diagnosticSink, "PPTX_UNSUPPORTED_CHART_NUMBER_FORMAT", OoxPdfSeverity.Warning, ChartNumberFormatConstructMessage(construct), chartPartName, slideIndex, "Ignored");
            }
        }

        if (sceneChart is not null)
        {
            foreach (PptxSceneChartAxis axis in sceneChart.Axes)
            {
                if (axis.NumberFormatInfo.IsDefined)
                {
                    collect(axis.NumberFormatInfo.FormatCode);
                }
            }

            foreach (PptxSceneChartPlot plot in sceneChart.Plots)
            {
                collect(plot.DataLabels.NumberFormatInfo.FormatCode);
                foreach (PptxSceneChartDataLabelOverride dataLabel in plot.DataLabels.Overrides)
                {
                    collect(dataLabel.NumberFormatInfo.FormatCode);
                }

                foreach (PptxSceneChartSeries series in plot.Series)
                {
                    collect(series.DataLabels.NumberFormatInfo.FormatCode);
                    foreach (PptxSceneChartDataLabelOverride dataLabel in series.DataLabels.Overrides)
                    {
                        collect(dataLabel.NumberFormatInfo.FormatCode);
                    }
                }
            }

            return;
        }

        foreach (XElement numberFormat in chartXml.Descendants(ChartNamespace + "numFmt"))
        {
            collect((string?)numberFormat.Attribute("formatCode"));
        }
    }

    private static string ChartNumberFormatConstructMessage(string construct)
    {
        return construct switch
        {
            "locale" => "Chart number formats use locale-specific rules that are not applied; values render with invariant formatting.",
            "color" => "Chart number formats specify colors that are not applied.",
            "native-digits" => "Chart number formats request native digit shapes that render as Western digits.",
            "bang-escape" => "Chart number formats use '!' escapes that are dropped.",
            "text-placeholder" => "Chart number formats use '@' text placeholders that are dropped.",
            _ => "Chart number formats contain unsupported bracketed constructs that are dropped.",
        };
    }

    // Locale decimal/group separators for calibrated LCIDs (S04; CLDR as independent spec).
    // System-locale ($-Fxxx) and unlisted LCIDs keep the "locale" diagnostic (open); locale
    // month/day names stay English (open). Table covers unambiguous decimal/group pairs only
    // (fr-FR narrow-nbsp and friends stay flagged).
    private static bool TryResolveChartNumberLocaleSeparators(string section, out string decimalSeparator, out string groupSeparator)
    {
        decimalSeparator = ".";
        groupSeparator = ",";
        bool inQuotes = false;
        for (int i = 0; i < section.Length; i++)
        {
            char c = section[i];
            if (!inQuotes && c == '\\' && i + 1 < section.Length)
            {
                i++;
                continue;
            }

            if (c == '"')
            {
                inQuotes = !inQuotes;
                continue;
            }

            if (inQuotes || c != '[')
            {
                continue;
            }

            int close = section.IndexOf(']', i);
            if (close < 0)
            {
                return false;
            }

            string inner = section.Substring(i + 1, close - i - 1).Trim();
            if (TryParseChartNumberLocaleId(inner, out int localeId)
                && TryGetChartNumberLocaleSeparators(localeId, out decimalSeparator, out groupSeparator))
            {
                return true;
            }
        }

        decimalSeparator = ".";
        groupSeparator = ",";
        return false;
    }

    private static bool TryParseChartNumberLocaleId(string inner, out int localeId)
    {
        localeId = 0;
        if (!inner.StartsWith('$'))
        {
            return false;
        }

        int dash = inner.LastIndexOf('-');
        if (dash < 1 || dash + 1 >= inner.Length)
        {
            return false;
        }

        string localeText = inner.Substring(dash + 1).Trim();
        if (!int.TryParse(localeText, System.Globalization.NumberStyles.HexNumber, CultureInfo.InvariantCulture, out localeId))
        {
            return false;
        }

        return localeId < 0xF000;
    }

    private static bool TryGetChartNumberLocaleSeparators(int localeId, out string decimalSeparator, out string groupSeparator)
    {
        switch (localeId)
        {
            case 0x0407:
            case 0x040A:
            case 0x0410:
            case 0x0413:
            case 0x0416:
                decimalSeparator = ",";
                groupSeparator = ".";
                return true;
            case 0x0409:
            case 0x0809:
            case 0x1009:
            case 0x0411:
                decimalSeparator = ".";
                groupSeparator = ",";
                return true;
            default:
                decimalSeparator = ".";
                groupSeparator = ",";
                return false;
        }
    }

    // [$currency-locale] symbol for numeric sections (S04); [$−locale] carries no
    // symbol. Locale separators and names stay invariant (open).
    private static string ParseChartNumberLocaleCurrency(string section)
    {
        bool inQuotes = false;
        for (int i = 0; i < section.Length; i++)
        {
            char c = section[i];
            if (!inQuotes && c == '\\' && i + 1 < section.Length)
            {
                i++;
                continue;
            }

            if (c == '"')
            {
                inQuotes = !inQuotes;
                continue;
            }

            if (inQuotes || c != '[')
            {
                continue;
            }

            int close = section.IndexOf(']', i);
            if (close < 0)
            {
                return string.Empty;
            }

            string inner = section.Substring(i + 1, close - i - 1).Trim();
            if (inner.StartsWith('$'))
            {
                int dash = inner.IndexOf('-', 1);
                return dash > 1 ? inner.Substring(1, dash - 1) : string.Empty;
            }
        }

        return string.Empty;
    }

    // Syntax view: escaped and quoted characters cannot act as format syntax. Escaped or
    // quoted percent/dollar keep distinct placeholders so they render literally without scaling.
    private static string ChartNumberFormatSyntaxView(string section)
    {
        var view = new StringBuilder(section.Length);
        bool inQuotes = false;
        for (int i = 0; i < section.Length; i++)
        {
            char c = section[i];
            if (c == '\\' && i + 1 < section.Length)
            {
                char escaped = section[++i];
                view.Append(escaped == '%' ? LiteralPercentPlaceholder : escaped == '$' ? LiteralDollarPlaceholder : '\0');
                continue;
            }

            if ((c == '_' || c == '*'))
            {
                // Width-cushion and fill runs carry no renderable value text;
                // mask the run so neighbours cannot misread the skipped char (S04).
                view.Append('\0');
                if (i + 1 < section.Length)
                {
                    view.Append('\0');
                    i++;
                }

                continue;
            }

            if (c == '"')
            {
                inQuotes = !inQuotes;
                view.Append('\0');
                continue;
            }

            if (!inQuotes && c == '[')
            {
                int bracketClose = section.IndexOf(']', i);
                if (bracketClose >= 0)
                {
                    view.Append('\0', bracketClose - i + 1);
                    i = bracketClose;
                    continue;
                }
            }

            view.Append(inQuotes ? (c == '%' ? LiteralPercentPlaceholder : c == '$' ? LiteralDollarPlaceholder : '\0') : c);
        }

        return view.ToString();
    }
}
