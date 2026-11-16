// Section parts: page setup, columns, headers/footers, paragraphs, runs and tables.
//
// 한글 keeps page setup (secPr) and the column definition (colPr) in the first run of a section's first
// paragraph; header and footer controls follow in the same run. Page and column breaks are paragraph attributes.

using System.Text;
using System.Xml.Linq;

namespace Filee.Engines.Hwp.Hwpx;

internal sealed partial class HwpxWriter
{
    // These names appear only in intermediate packages for rhwp.
    internal const string FirstPageHeaderName = "FileeFirstPageHeader";
    internal const string PageFieldMarker = "filee-page-field";

    /// <summary>Width available to the content being written (text width, column, table cell or text box).</summary>
    private int _width = 42520;

    /// <summary>Distance from the paper edge to the text area of the current section (left, and top including the header).</summary>
    private int _pageLeft = 7200;
    private int _pageTop = 8505;
    private int _paperWidth = 59530;

    /// <summary>Style of plain paragraphs in the current container (body text, or footnote text inside notes).</summary>
    private (string StyleId, string ParaPrId, string CharPrId) _base;

    // ───────────────────────── Sections ─────────────────────────

    private string Section(HSection section)
    {
        var secPr = new XElement(_templateSection.Descendants(Hp + "secPr").FirstOrDefault()
                                 ?? throw new InvalidDataException("The template section has no secPr."));
        var pagePr = secPr.Element(Hp + "pagePr");
        var margin = pagePr?.Element(Hp + "margin");
        if (section.Page is { } page && pagePr is not null && margin is not null)
        {
            // Width/height always describe the portrait sheet; "NARROWLY" turns it (landscape), "WIDELY" does not.
            pagePr.SetAttributeValue("landscape", page.Landscape ? "NARROWLY" : "WIDELY");
            pagePr.SetAttributeValue("width", Math.Min(page.Width, page.Height));
            pagePr.SetAttributeValue("height", Math.Max(page.Width, page.Height));
            margin.SetAttributeValue("left", page.Left);
            margin.SetAttributeValue("right", page.Right);
            margin.SetAttributeValue("top", page.Top);
            margin.SetAttributeValue("bottom", page.Bottom);
            margin.SetAttributeValue("header", page.Header);
            margin.SetAttributeValue("footer", page.Footer);
            margin.SetAttributeValue("gutter", page.Gutter);
        }
        if (secPr.Element(Hp + "visibility") is { } visibility)
        {
            visibility.SetAttributeValue("hideFirstHeader", section.HideFirstHeader || section.FirstPageHeader is not null ? 1 : 0);
            visibility.SetAttributeValue("hideFirstFooter", section.HideFirstFooter ? 1 : 0);
        }
        secPr.Element(Hp + "startNum")?.SetAttributeValue("page", section.StartPageNumber ?? 0);

        int Margin(string name, int fallback) => (int?)margin?.Attribute(name) ?? fallback;
        var paperWidth = (int?)pagePr?.Attribute("width") ?? 59530;
        var paperHeight = (int?)pagePr?.Attribute("height") ?? 84190;
        var landscape = (string?)pagePr?.Attribute("landscape") == "NARROWLY";
        _paperWidth = landscape ? paperHeight : paperWidth;
        var textWidth = Math.Max(2000, _paperWidth - Margin("left", 7200) - Margin("right", 7200) - Margin("gutter", 0));
        _base = (_normalStyleId, _normalParaPrId, _normalCharPrId);

        // The header area is as tall as the header, as in Word. rhwp also stops at the end of the area, which left out a
        // table after an empty paragraph in a header with little room above the body (issue #1152).
        if (section.HeaderMovesBody && section.Headers.Count > 0)
            margin?.SetAttributeValue("header", Math.Max(Margin("header", 4250), section.Headers.Max(h => EstimateHeight(h.Blocks, textWidth))));
        // The footer grows inward from the Word footer distance. A common footer can reserve its height for the
        // whole section. Parity-specific or delayed footers would need different body sizes on different pages.
        if (_forRhwpPdf && section.FooterMovesBody
            && section.Footers.FirstOrDefault(f => f is { Pages: HPageType.Both, StartsAt: null }) is { } commonFooter)
            margin?.SetAttributeValue("footer", Math.Max(Margin("footer", 2240), EstimateHeight(commonFooter.Blocks, textWidth)));
        _pageLeft = Margin("left", 7200) + Margin("gutter", 0);
        _pageTop = Margin("top", 4255) + Margin("header", 4250);

        var setup = new StringBuilder();
        setup.Append(Inner(secPr));
        setup.Append("<hp:ctrl>").Append(ColumnDefinition(section.Columns)).Append("</hp:ctrl>");
        // rhwp goes by this control, not by startNum, for a section that restarts its page numbers.
        if (section.StartPageNumber is { } startPage)
            setup.Append($"<hp:ctrl><hp:newNum num=\"{startPage}\" numType=\"PAGE\"/></hp:ctrl>");

        _width = textWidth;
        var controlId = 1;
        var later = new Dictionary<HBlock, string>();
        void Add(HHeaderFooter item, string control)
        {
            if (item.StartsAt is { } paragraph)
                later[paragraph] = later.GetValueOrDefault(paragraph) + control;
            else
                setup.Append(control);
        }
        foreach (var header in section.Headers)
            Add(header, HeaderFooter("header", header, controlId++, textWidth, Margin("header", 4250), "TOP"));
        foreach (var footer in section.Footers)
            Add(footer, HeaderFooter("footer", footer, controlId++, textWidth, Margin("footer", 2240), "BOTTOM"));

        _width = ColumnWidth(textWidth, section.Columns);

        // 한글 has no header for the first page only: it is a text box at the header's place, anchored in the section's
        // first paragraph, and the other header is hidden there. The text below goes around it, as Word moves the body.
        if (section.FirstPageHeader is { } firstPageHeader)
        {
            var firstBlock = 0;
            var paddingTop = 0;
            if (_forRhwpPdf)
            {
                // rhwp collapses leading empty paragraphs when it draws a text box's table. Preserve the space
                // as padding while retaining the box's full height so the body starts in the same place.
                var blank = firstPageHeader.TakeWhile(b => b is HParagraph { Inlines.Count: 0, List: null } p
                    && (p.Format.Borders is null || p.Format.Borders == HBorders.Empty)
                    && p.Format.Fill is null && p.Format.LeftBar is null).ToList();
                if (blank.Count < firstPageHeader.Count && firstPageHeader[blank.Count] is HTable)
                {
                    firstBlock = blank.Count;
                    paddingTop = Math.Max(0, EstimateHeight(blank, textWidth));
                }
            }
            var box = new HTextBox
            {
                Width = textWidth,
                Height = EstimateHeight(firstPageHeader, textWidth),
                Anchor = new HAnchor("PAPER", "LEFT", _pageLeft, "PAPER", "TOP", Margin("top", 4255), "TOP_AND_BOTTOM", true),
                Line = HBorder.None,
                Padding = new HInsets(0, 0, paddingTop, 0),
            };
            box.Blocks.AddRange(firstPageHeader.Skip(firstBlock));
            setup.Append(TextBox(box, _forRhwpPdf ? FirstPageHeaderName : ""));
        }

        var body = new StringBuilder(SectionOpenTag());
        var blocks = section.Blocks.Count == 0 || section.Blocks[0] is not HParagraph
            ? section.Blocks.Prepend(new HParagraph()).ToList()
            : section.Blocks;
        var wrappedLines = new Dictionary<int, string>();
        for (var i = 0; i < blocks.Count; i++)
        {
            if (blocks[i] is HParagraph { ColumnsChange: { } columns })
                _width = ColumnWidth(textWidth, columns);
            var controls = (i == 0 ? setup.ToString() : "") + later.GetValueOrDefault(blocks[i]);
            string? leadingRun = null;
            if (controls.Length > 0)
            {
                var controlCharPr = _normalCharPrId;
                // rhwp uses the run font as the minimum height of an otherwise empty setup paragraph.
                if (_forRhwpPdf && blocks[i] is HParagraph
                    { Inlines.Count: 0, Format.LineSpacing: { Kind: HLineSpacingKind.Fixed, Value: > 0 and var height } }
                    && height < BaseSize(_normalCharPrId))
                    controlCharPr = CharPr(_normalCharPrId, new HCharFormat(Size: height));
                leadingRun = $"<hp:run charPrIDRef=\"{controlCharPr}\">{controls}</hp:run>";
            }
            var xml = blocks[i] switch
            {
                HTable { Anchor: not null } table when _forRhwpPdf =>
                    FloatingTable(table, leadingRun, blocks, i, wrappedLines),
                HParagraph p when wrappedLines.Remove(i, out var line) => Paragraph(p, leadingRun, line),
                HParagraph p when _forRhwpPdf && PicturesHaveClearSpace(p, blocks, i) =>
                    Paragraph(p, leadingRun, floatPicturesInFront: true),
                HParagraph p when _forRhwpPdf => Paragraph(p, leadingRun, drawEquationBrace: true,
                    preserveEmptyLine: i == 0 && section.FirstPageHeader is not null && section.Blocks.FirstOrDefault() == p),
                var block => Block(block, leadingRun),
            };
            body.Append(xml);
        }
        return body.Append("</hs:sec>").ToString();
    }

    private static int ColumnWidth(int textWidth, HColumns columns) =>
        columns.Count <= 1 ? textWidth : Math.Max(2000, (textWidth - columns.Gap * (columns.Count - 1)) / columns.Count);

    private static string ColumnDefinition(HColumns columns)
    {
        var line = columns.Separator ? "<hp:colLine type=\"SOLID\" width=\"0.12 mm\" color=\"#000000\"/>" : "";
        return $"<hp:colPr id=\"\" type=\"NEWSPAPER\" layout=\"LEFT\" colCount=\"{Math.Max(1, columns.Count)}\" sameSz=\"1\" sameGap=\"{(columns.Count > 1 ? columns.Gap : 0)}\">{line}</hp:colPr>";
    }

    private string HeaderFooter(string kind, HHeaderFooter item, int id, int textWidth, int textHeight, string verticalAlign)
    {
        var pages = item.Pages switch
        {
            HPageType.Even => "EVEN",
            HPageType.Odd => "ODD",
            _ => "BOTH",
        };
        return $"<hp:ctrl><hp:{kind} id=\"{id}\" applyPageType=\"{pages}\">" +
               SubList(item.Blocks, verticalAlign, textWidth, textHeight) +
               $"</hp:{kind}></hp:ctrl>";
    }

    /// <summary>Paragraph list of a header, footer, cell, text box or note.</summary>
    private string SubList(List<HBlock> blocks, string verticalAlign, int textWidth = 0, int textHeight = 0, string? leadingRun = null)
    {
        var sb = new StringBuilder();
        sb.Append($"<hp:subList id=\"\" textDirection=\"HORIZONTAL\" lineWrap=\"BREAK\" vertAlign=\"{verticalAlign}\" linkListIDRef=\"0\" linkListNextIDRef=\"0\" textWidth=\"{textWidth}\" textHeight=\"{textHeight}\" hasTextRef=\"0\" hasNumRef=\"0\">");
        var list = blocks.Count == 0 || (leadingRun is not null && blocks[0] is not HParagraph) ? blocks.Prepend(new HParagraph()).ToList() : blocks;
        for (var i = 0; i < list.Count; i++)
            sb.Append(Block(list[i], i == 0 ? leadingRun : null));
        return sb.Append("</hp:subList>").ToString();
    }

    // ───────────────────────── Blocks ─────────────────────────

    /// <param name="leadingRun">Raw run placed first in the paragraph (section setup, note number).</param>
    private string Block(HBlock block, string? leadingRun = null) => block switch
    {
        HParagraph paragraph => Paragraph(paragraph, leadingRun),
        HTable table => Table(table, leadingRun),
        _ => "",
    };

    private string Paragraph(HParagraph paragraph, string? leadingRun, string? wrappedLine = null,
        bool floatPicturesInFront = false, bool drawEquationBrace = false, bool preserveEmptyLine = false)
    {
        _plainText.AppendLine();
        var (styleId, paraPrBase, charPrBase) = paragraph.HeadingLevel > 0 && _headings.TryGetValue(paragraph.HeadingLevel, out var heading)
            ? heading
            : _base;
        var paraPrId = ParaPr(paraPrBase, paragraph.Format, paragraph.List);

        var runs = new RunBuilder();
        if (paragraph.ColumnsChange is { } columns)
            runs.Control(_normalCharPrId, $"<hp:ctrl>{ColumnDefinition(columns)}</hp:ctrl>");
        var size = LargestSize(paragraph.Inlines) ?? paragraph.MarkFormat.Size ?? BaseSize(charPrBase);
        var tabs = new Queue<string>(TabElements(paragraph, size));
        var equationLine = _forRhwpPdf ? EquationLine(paragraph, charPrBase, drawEquationBrace) : null;
        var markCharPr = CharPr(charPrBase, paragraph.MarkFormat);
        Inlines(paragraph.Inlines, runs, charPrBase, link: false, tabs, equationLine?.Height,
            floatPicturesInFront, floatPicturesInFront && paragraph.Inlines.All(i => i is HImage) ? markCharPr : null,
            equationLine?.Brace);
        // rhwp skips the line of an empty paragraph carrying the first-page header's floating box.
        if (preserveEmptyLine && paragraph.Inlines.Count == 0 && paragraph.List is null)
            runs.Text(markCharPr, "&#160;");

        var sb = new StringBuilder();
        sb.Append($"<hp:p id=\"0\" paraPrIDRef=\"{paraPrId}\" styleIDRef=\"{styleId}\" pageBreak=\"{(paragraph.PageBreakBefore ? 1 : 0)}\" columnBreak=\"{(paragraph.ColumnBreakBefore ? 1 : 0)}\" merged=\"0\">");
        sb.Append(leadingRun);
        sb.Append(runs.Finish(markCharPr, leadingRun is not null));
        if (equationLine is { } line)
            sb.Append(line.Xml);
        else
            sb.Append(wrappedLine);
        return sb.Append("</hp:p>").ToString();
    }

    /// <summary>
    /// Collects runs: consecutive content with the same character shape shares one hp:run and one hp:t, and
    /// controls (fields, bookmarks) and objects (pictures, tables, text boxes) go between the text parts.
    /// </summary>
    private sealed class RunBuilder
    {
        private readonly StringBuilder _output = new();
        private readonly StringBuilder _items = new();
        private string? _charPr;
        private bool _textOpen;
        private bool _endsWithObject;

        public void Text(string charPr, string xml)
        {
            Switch(charPr);
            if (!_textOpen)
            {
                _items.Append("<hp:t>");
                _textOpen = true;
            }
            _items.Append(xml);
            _endsWithObject = false;
        }

        public void Control(string charPr, string xml, bool isObject = false)
        {
            Switch(charPr);
            CloseText();
            _items.Append(xml);
            _endsWithObject = isObject;
        }

        /// <summary>The runs; an empty paragraph still needs one run for 한글 to show it.</summary>
        public string Finish(string emptyCharPr, bool hasLeadingRun)
        {
            Flush();
            return _output.Length == 0 && !hasLeadingRun ? $"<hp:run charPrIDRef=\"{emptyCharPr}\"><hp:t/></hp:run>" : _output.ToString();
        }

        private void Switch(string charPr)
        {
            if (_charPr == charPr)
                return;
            Flush();
            _charPr = charPr;
        }

        private void CloseText()
        {
            if (_textOpen)
                _items.Append("</hp:t>");
            _textOpen = false;
        }

        private void Flush()
        {
            if (_charPr is null)
                return;
            CloseText();
            if (_endsWithObject)
                _items.Append("<hp:t/>"); // as 한글 writes it after an object
            _output.Append("<hp:run charPrIDRef=\"").Append(_charPr).Append("\">").Append(_items).Append("</hp:run>");
            _items.Clear();
            _charPr = null;
            _endsWithObject = false;
        }
    }

    /// <param name="tabs">Pre-computed inline tab elements of the paragraph, in order (see TabElements).</param>
    private void Inlines(IEnumerable<HInline> inlines, RunBuilder runs, string charPrBase, bool link, Queue<string> tabs,
        int? equationLineHeight = null, bool floatPicturesInFront = false, string? pictureCharPr = null, string? equationBrace = null)
    {
        foreach (var inline in inlines)
        {
            switch (inline)
            {
                case HEquation equation:
                    runs.Control(CharPr(charPrBase, equation.Format, link),
                        Equation(equation, charPrBase, equationBrace is not null) + equationBrace, isObject: true);
                    _plainText.Append(equation.Text);
                    break;
                case HCachedPageField cached when _forRhwpPdf:
                    runs.Control(CharPr(charPrBase, cached.Format, link),
                        $"<?{PageFieldMarker} {cached.Kind}?><hp:t>{Escape(cached.Text)}</hp:t>");
                    _plainText.Append(cached.Text);
                    break;
                case HText text:
                    runs.Text(CharPr(charPrBase, text.Format, link), Escape(text.Text));
                    _plainText.Append(text.Text);
                    break;
                case HLineBreak lineBreak:
                    runs.Text(CharPr(charPrBase, lineBreak.Format, link), "<hp:lineBreak/>");
                    _plainText.AppendLine();
                    break;
                case HTab tab:
                    runs.Text(CharPr(charPrBase, tab.Format, link), tabs.Count > 0 ? tabs.Dequeue() : "<hp:tab width=\"4000\" leader=\"0\" type=\"1\"/>");
                    _plainText.Append('\t');
                    break;
                case HLink hyperlink:
                    {
                        var id = NextId();
                        var format = hyperlink.Content.OfType<HText>().FirstOrDefault()?.Format ?? default;
                        var charPr = CharPr(charPrBase, format, link: true);
                        runs.Control(charPr, FieldBegin(id, hyperlink.Target));
                        Inlines(hyperlink.Content, runs, charPrBase, link: true, tabs);
                        runs.Control(charPr, $"<hp:ctrl><hp:fieldEnd beginIDRef=\"{id}\" fieldid=\"{HyperlinkFieldId}\"/></hp:ctrl>");
                        break;
                    }
                case HBookmark bookmark:
                    runs.Control(charPrBase, $"<hp:ctrl><hp:bookmark name=\"{Escape(bookmark.Name)}\"/></hp:ctrl>");
                    break;
                case HField field:
                    runs.Control(CharPr(charPrBase, field.Format), AutoNumber(field.Kind == HFieldKind.PageNumber ? "PAGE" : "TOTAL_PAGE", 1, ""));
                    break;
                case HNote note:
                    runs.Control(charPrBase, Note(note));
                    break;
                case HImage image:
                    if (Picture(image, floatPicturesInFront) is { } picture)
                        runs.Control(pictureCharPr ?? charPrBase, picture, isObject: true);
                    break;
                case HTextBox box:
                    var anchor = box.IsEquationNumber && box.Anchor is { } original && equationLineHeight is { } height
                        ? original with { VerticalOffset = Math.Max(0, (height - box.Height) / 2) }
                        : box.Anchor;
                    runs.Control(charPrBase, TextBox(box, anchorOverride: anchor), isObject: true);
                    break;
                case HShape shape:
                    runs.Control(charPrBase, Shape(shape), isObject: true);
                    break;
            }
        }
    }

    // ───────────────────────── Tables ─────────────────────────

    private string Table(HTable table, string? leadingRun, HAnchor? anchorOverride = null, string? wrappedLine = null)
    {
        if (table.Rows.Count == 0)
            return leadingRun is null ? "" : Paragraph(new HParagraph(), leadingRun);

        var (placed, columnCount) = PlaceCells(table);

        var widths = ColumnWidths(table, columnCount, TableAvailableWidth(table));
        var totalWidth = widths.Sum();
        var rowHeights = EstimateRowHeights(table, placed, widths);
        var margin = table.CellMargin ?? new HInsets(510, 510, 141, 141);
        var tableBorders = table.Borders ?? HBorders.Grid;
        var headerRows = table.Rows.TakeWhile(r => r.Header).Count();
        var anchor = anchorOverride ?? table.Anchor;

        var sb = new StringBuilder();
        var id = NextId();
        sb.Append($"<hp:tbl id=\"{id}\" zOrder=\"0\" numberingType=\"TABLE\" textWrap=\"{anchor?.Wrap ?? "TOP_AND_BOTTOM"}\" textFlow=\"BOTH_SIDES\" lock=\"0\" dropcapstyle=\"None\" pageBreak=\"CELL\" repeatHeader=\"{(headerRows > 0 ? 1 : 0)}\" rowCnt=\"{table.Rows.Count}\" colCnt=\"{columnCount}\" cellSpacing=\"0\" borderFillIDRef=\"{BorderFillId(tableBorders, null)}\" noAdjust=\"0\">");
        sb.Append($"<hp:sz width=\"{totalWidth}\" widthRelTo=\"ABSOLUTE\" height=\"{rowHeights.Sum()}\" heightRelTo=\"ABSOLUTE\" protect=\"0\"/>");
        var horizontalAlign = table.Align switch
        {
            HAlign.Center => "CENTER",
            HAlign.Right => "RIGHT",
            _ => "LEFT",
        };
        sb.Append(Position(anchor ?? new HAnchor("COLUMN", horizontalAlign, table.Align == HAlign.Left ? table.Indent : 0,
            "PARA", "TOP", 0, "TOP_AND_BOTTOM", false)));
        var clearance = anchor is null ? new HInsets(0, 0, 0, table.Tight ? 0 : 1417) : table.OuterMargin;
        sb.Append($"<hp:outMargin left=\"{clearance.Left}\" right=\"{clearance.Right}\" top=\"{clearance.Top}\" bottom=\"{clearance.Bottom}\"/>");
        sb.Append($"<hp:inMargin left=\"{margin.Left}\" right=\"{margin.Right}\" top=\"{margin.Top}\" bottom=\"{margin.Bottom}\"/>");

        var outerWidth = _width;
        for (var r = 0; r < placed.Count; r++)
        {
            sb.Append("<hp:tr>");
            foreach (var (cell, column) in placed[r])
            {
                var cellMargin = cell.CellMargin ?? margin;
                var width = 0;
                for (var i = column; i < column + cell.ColSpan && i < widths.Length; i++)
                    width += widths[i];
                var height = 0;
                for (var i = r; i < r + cell.RowSpan && i < rowHeights.Length; i++)
                    height += rowHeights[i];

                if (_forRhwpPdf)
                    cellMargin = RenderCellMargin(table, cell, cellMargin, width);
                _width = Math.Max(1000, width - cellMargin.Left - cellMargin.Right);
                var verticalAlign = cell.VerticalAlign switch
                {
                    HVerticalAlign.Center => "CENTER",
                    HVerticalAlign.Bottom => "BOTTOM",
                    _ => "TOP",
                };
                sb.Append($"<hp:tc name=\"\" header=\"{(r < headerRows ? 1 : 0)}\" hasMargin=\"{(cell.CellMargin is not null || cellMargin != margin ? 1 : 0)}\" protect=\"0\" editable=\"0\" dirty=\"0\" borderFillIDRef=\"{BorderFillId(cell.Borders ?? tableBorders, cell.Fill)}\">");
                sb.Append(SubList(cell.Blocks, verticalAlign));
                sb.Append($"<hp:cellAddr colAddr=\"{column}\" rowAddr=\"{r}\"/>");
                sb.Append($"<hp:cellSpan colSpan=\"{cell.ColSpan}\" rowSpan=\"{cell.RowSpan}\"/>");
                // Like 한글 itself: a merged cell is as tall as all the rows it spans.
                sb.Append($"<hp:cellSz width=\"{width}\" height=\"{height}\"/>");
                sb.Append($"<hp:cellMargin left=\"{cellMargin.Left}\" right=\"{cellMargin.Right}\" top=\"{cellMargin.Top}\" bottom=\"{cellMargin.Bottom}\"/>");
                sb.Append("</hp:tc>");
            }
            sb.Append("</hp:tr>");
        }
        _width = outerWidth;
        sb.Append("</hp:tbl>");

        // The paragraph holding the table has a text line under it; a tight table's is 1 pt with no spacing.
        var (paraPrId, charPrId) = table.Tight
            ? (ParaPr(_base.ParaPrId, new HParaFormat(Before: 0, After: 0, LineSpacing: new HLineSpacing(HLineSpacingKind.Percent, 100)), null),
               CharPr(_base.CharPrId, new HCharFormat(Size: 100)))
            : (_base.ParaPrId, _base.CharPrId);
        var runs = new RunBuilder();
        runs.Control(charPrId, sb.ToString(), isObject: true);
        var paragraph = $"<hp:p id=\"0\" paraPrIDRef=\"{paraPrId}\" styleIDRef=\"{_base.StyleId}\" pageBreak=\"{(table.PageBreakBefore ? 1 : 0)}\" columnBreak=\"0\" merged=\"0\">{leadingRun}{runs.Finish(charPrId, leadingRun is not null)}{wrappedLine}</hp:p>";
        return paragraph + string.Concat(table.Caption.Select(b => Block(b)));
    }

    // Absolute column widths of a table anchored to the paper can extend into the body margins.
    private int TableAvailableWidth(HTable table) => Math.Max(4000,
        table.ColumnWidths is { Length: > 0 } && table.Anchor is
            { HorizontalRelativeTo: "PAPER", HorizontalAlign: "LEFT", HorizontalOffset: var x }
            ? _paperWidth - x
            : _width - table.Indent);

    /// <summary>
    /// Grid position of every cell: explicit cell columns (DOCX) or the next free position (Pandoc), skipping
    /// positions covered by cells spanning rows from above.
    /// </summary>
    internal static (List<List<(HCell Cell, int Column)>> Placed, int Columns) PlaceCells(HTable table)
    {
        var placed = new List<List<(HCell, int)>>();
        var occupied = new HashSet<(int Row, int Column)>();
        var columnCount = table.ColumnCount;
        for (var r = 0; r < table.Rows.Count; r++)
        {
            var row = new List<(HCell, int)>();
            var column = 0;
            foreach (var cell in table.Rows[r].Cells)
            {
                if (cell.Column >= 0)
                    column = cell.Column;
                while (occupied.Contains((r, column)))
                    column++;
                for (var dr = 0; dr < cell.RowSpan; dr++)
                {
                    for (var dc = 0; dc < cell.ColSpan; dc++)
                        occupied.Add((r + dr, column + dc));
                }
                row.Add((cell, column));
                column += cell.ColSpan;
                columnCount = Math.Max(columnCount, column);
            }
            placed.Add(row);
        }
        return (placed, Math.Max(1, columnCount));
    }

    /// <summary>Column widths that fit <paramref name="available"/>: absolute (scaled down if too wide), relative or equal.</summary>
    internal static int[] ColumnWidths(HTable table, int columns, int available)
    {
        if (table.ColumnWidths is { Length: > 0 } absolute)
        {
            var widths = Enumerable.Range(0, columns).Select(i => i < absolute.Length ? absolute[i] : (int)absolute.Average()).ToArray();
            var total = widths.Sum();
            return total <= available ? widths : widths.Select(w => (int)((long)w * available / total)).ToArray();
        }

        var fractions = Enumerable.Range(0, columns).Select(i => table.RelativeWidths is { } rel && i < rel.Length ? rel[i] : 0).ToArray();
        var known = fractions.Where(f => f > 0).Sum();
        var unknown = fractions.Count(f => f <= 0);
        var scale = known > 1 ? 1 / known : 1; // never exceed the text width
        var rest = unknown == 0 ? 0 : Math.Max(0, 1 - known * scale) / unknown;
        if (unknown > 0 && rest <= 0)
            return Enumerable.Repeat(available / columns, columns).ToArray(); // explicit columns used all the width
        return fractions.Select(f => (int)(available * (f > 0 ? f * scale : rest))).ToArray();
    }

    /// <summary>
    /// Serializes an element with the usual prefixes (hp:, hc: ...) but without namespace declarations, which the
    /// section root already makes. A detached element would otherwise be written with a default namespace
    /// (&lt;secPr xmlns="..."&gt;), which 한글 and other HWPX readers that match prefixed names do not accept.
    /// </summary>
    private static string Inner(XElement element)
    {
        var copy = new XElement(element);
        foreach (var (prefix, ns) in new[] { ("hp", Hp), ("hc", Hc), ("hh", Hh), ("hs", Hs) })
            copy.SetAttributeValue(XNamespace.Xmlns + prefix, ns.NamespaceName);
        var xml = copy.ToString(SaveOptions.DisableFormatting);
        foreach (var (prefix, ns) in new[] { ("hp", Hp), ("hc", Hc), ("hh", Hh), ("hs", Hs) })
            xml = xml.Replace($" xmlns:{prefix}=\"{ns.NamespaceName}\"", "", StringComparison.Ordinal);
        return xml;
    }
}
