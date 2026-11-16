using SkiaSharp;

namespace Filee.Engines.Hwp.Hwpx;

internal sealed partial class HwpxWriter
{
    /// <summary>rhwp's table wrap needs the saved positions of the lines beside the table.</summary>
    private string FloatingTable(HTable table, string? leadingRun, IReadOnlyList<HBlock> blocks, int index,
        Dictionary<int, string> wrappedLines)
    {
        var anchor = table.Anchor!;
        if (anchor is not { VerticalRelativeTo: "PARA", VerticalAlign: "TOP" } || table.Rows.Count == 0)
            return Table(table, leadingRun);

        var (cells, columns) = PlaceCells(table);
        var widths = ColumnWidths(table, columns, Math.Max(4000, _width - table.Indent));
        var height = EstimateRowHeights(table, cells, widths).Sum();
        var space = table.OuterMargin;
        if (anchor is { HorizontalAlign: "CENTER", HorizontalOffset: 0 })
        {
            var blankSpace = 0;
            for (var i = index + 1; i < blocks.Count && blocks[i] is HParagraph { Inlines.Count: 0 } p && SameFlow(p); i++)
            {
                blankSpace += EstimateHeight(p, _width);
                if (blankSpace >= anchor.VerticalOffset + height + space.Bottom)
                    return Table(table, leadingRun, anchor with { Wrap = "BEHIND_TEXT" });
            }
        }

        if (anchor is not { HorizontalRelativeTo: "COLUMN" or "PAGE", HorizontalAlign: "LEFT" or "RIGHT" })
            return Table(table, leadingRun);
        var tableWidth = widths.Sum();
        var wrapsRight = anchor.HorizontalAlign == "LEFT";
        var left = wrapsRight ? anchor.HorizontalOffset + tableWidth + space.Right : 0;
        var width = wrapsRight ? _width - left : _width - tableWidth - space.Left;
        if (width < 1000)
            return Table(table, leadingRun);

        var bodyTop = 0;
        var bottom = anchor.VerticalOffset + space.Top + height + space.Bottom;
        var added = false;
        for (var i = index + 1; i < blocks.Count && blocks[i] is HParagraph p && SameFlow(p); i++)
        {
            var trailingBreak = p.Inlines.LastOrDefault() is HLineBreak;
            var text = p.Inlines.Take(p.Inlines.Count - (trailingBreak ? 1 : 0)).OfType<HText>().ToList();
            if (text.Count == 0 || text.Count != p.Inlines.Count - (trailingBreak ? 1 : 0) ||
                p.Format is { RightToLeft: true } or { FirstLine: not (null or 0) })
                break;
            var size = LargestSize(p.Inlines) ?? p.MarkFormat.Size ?? BaseSize(_base.CharPrId);
            var before = p.Format.Before ?? 0;
            var after = p.Format.After ?? 0;
            var lineHeight = (EstimateHeight(p, _width) - before - after) / (trailingBreak ? 2 : 1);
            var available = width - (p.Format.Left ?? 0) - (p.Format.Right ?? 0);
            var required = text.Sum(run => EstimateRunWidth(run, run.Format.Size ?? size));
            var top = bodyTop + before;
            if (top + size <= anchor.VerticalOffset - space.Top || top >= bottom || required > available)
                break;

            var lines = StoredLine(top, size, lineHeight - size, left, width);
            if (trailingBreak)
            {
                // The empty line after a manual break can return to full width below the table.
                var markSize = p.MarkFormat.Size ?? size;
                lines += StoredLine(Math.Max(bottom, top + lineHeight), markSize, 0, 0, _width, text.Sum(run => run.Text.Length) + 1);
            }
            wrappedLines[i] = $"<hp:linesegarray>{lines}</hp:linesegarray>";
            added = true;
            bodyTop = top + lineHeight + after;
            if (trailingBreak)
                break;
        }

        var hostLine = added ? $"<hp:linesegarray>{StoredLine(0, 100, 0, left, width)}</hp:linesegarray>" : null;
        return Table(table, leadingRun, wrappedLine: hostLine);
    }

    private static string StoredLine(int top, int height, int spacing, int left, int width, int textPosition = 0) =>
        $"<hp:lineseg textpos=\"{textPosition}\" vertpos=\"{top}\" vertsize=\"{height}\" textheight=\"{height}\" " +
        $"baseline=\"{height * 85 / 100}\" spacing=\"{spacing}\" horzpos=\"{left}\" horzsize=\"{width}\" flags=\"393216\"/>";

    /// <summary>
    /// rhwp can rewind body flow after a behind-text picture anchored to the paragraph. When the drawing
    /// fits in the empty space under the current line, drawing it in front preserves both the ink and the flow.
    /// </summary>
    private bool PicturesHaveClearSpace(HParagraph paragraph, IReadOnlyList<HBlock> blocks, int index)
    {
        if (paragraph.Inlines.Any(item => item is not (HText or HImage)))
            return false;
        var pictures = paragraph.Inlines.OfType<HImage>().ToList();
        if (pictures.Count == 0 || pictures.Any(image => image.Height is null || image.Anchor is not
            { VerticalRelativeTo: "PARA", VerticalAlign: "TOP", Wrap: "BEHIND_TEXT" }))
            return false;

        var size = LargestSize(paragraph.Inlines) ?? paragraph.MarkFormat.Size ?? BaseSize(_base.CharPrId);
        if (paragraph.Inlines.OfType<HText>().Any(text => !string.IsNullOrWhiteSpace(text.Text)) &&
            pictures.Any(image => image.Anchor!.VerticalOffset < (paragraph.Format.Before ?? 0) + size))
            return false;

        var bottom = pictures.Max(image => (long)image.Anchor!.VerticalOffset + image.Height!.Value);
        var clear = EstimateHeight(paragraph, _width);
        for (var i = index + 1; i < blocks.Count && blocks[i] is HParagraph p && SameFlow(p); i++)
        {
            if (p.Inlines.Any(inline => inline is not HImage { Anchor: not null } &&
                (inline is not HText run || !string.IsNullOrWhiteSpace(run.Text))))
                break;
            clear += EstimateHeight(p, _width);
            if (clear >= bottom)
                return true;
        }
        return clear >= bottom;
    }

    private static bool SameFlow(HParagraph paragraph) =>
        !paragraph.PageBreakBefore && !paragraph.ColumnBreakBefore && paragraph.ColumnsChange is null && paragraph.List is null;

    private HInsets RenderCellMargin(HTable table, HCell cell, HInsets margin, int width)
    {
        if (table.Anchor is not null && cell.VerticalAlign == HVerticalAlign.Top &&
            cell.Blocks.FirstOrDefault() is HParagraph { Format.Before: > 0 and var before } first)
        {
            // In a floating table, rhwp draws the first line at the cell padding. Carry the Word
            // paragraph's leading space there, including the extra above the first glyph baseline.
            var size = LargestSize(first.Inlines) ?? first.MarkFormat.Size ?? BaseSize(_base.CharPrId);
            var leading = Math.Max(0, LineHeight(first, size, _base.ParaPrId) - size);
            margin = margin with { Top = margin.Top + before + leading };
        }

        if (!table.Tight || !table.WordAutoFit || cell.CellMargin is not null || cell.Blocks is not
            [HParagraph { Inlines: [HText text] } paragraph] || text.GetType() != typeof(HText) ||
            text.Format is not { Font: { } face, Size: > 0 and var fontSize, Italic: not true, Spacing: null or 0 } ||
            text.Text.Length < 2 || !text.Text.All(char.IsAsciiLetter))
            return margin;

        var available = width - margin.Left - margin.Right - (paragraph.Format.Left ?? 0) -
                        (paragraph.Format.Right ?? 0) - (paragraph.Format.FirstLine ?? 0);
        var approximate = TextWidth(text.Text, fontSize);
        if (approximate < available * 0.7 || approximate > available * 1.3)
            return margin;
        var typeface = SKFontManager.Default.MatchFamily(face, text.Format.Bold == true ? SKFontStyle.Bold : SKFontStyle.Normal);
        if (typeface is null || !typeface.FamilyName.Equals(face, StringComparison.OrdinalIgnoreCase))
            return margin;
        using var font = new SKFont(typeface, fontSize);
        var shortfall = (int)Math.Ceiling(font.MeasureText(text.Text)) - available;
        // Word's automatic layout can fit a word right against the padding. Allow for
        // a small metric difference without changing the grid or the position of its borders.
        var adjustment = (shortfall + 11) / 2;
        return shortfall is > 0 and <= 100 && Math.Min(margin.Left, margin.Right) >= adjustment
            ? margin with { Left = margin.Left - adjustment, Right = margin.Right - adjustment }
            : margin;
    }
}
