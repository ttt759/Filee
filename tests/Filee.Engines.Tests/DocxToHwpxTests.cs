// DOCX → HWPX with the built-in reader and writer: layout that Pandoc-based conversion loses (page setup,
// headers/footers, columns, sections, text boxes, shapes, floating pictures, formatting, table borders) must
// arrive in the OWPML structure 한글 uses. Every file is also checked for package validity and rendered by rhwp.

using System.Xml.Linq;
using Filee.Core.Conversion;
using Filee.Core.Presets;
using Filee.Engines.Hwp.Hwpx;
using Filee.Engines.Hwp.Hwpx.Docx;
using ImageMagick;
using static Filee.Engines.Tests.DocxBuilder;
using static Filee.Engines.Tests.HwpxAssert;

namespace Filee.Engines.Tests;

public class DocxToHwpxTests(EngineFixture fx) : IClassFixture<EngineFixture>
{
    private async Task<string> ConvertAsync(DocxBuilder docx, string name, int? expectImages = 0)
    {
        var path = docx.Save(Path.Combine(fx.NewFolder(), name + ".docx"));
        var job = await fx.ConvertAsync([path], new Preset { TargetFormat = "hwpx" });
        Assert.True(job.State == JobState.Completed, string.Join("; ", job.Files.Select(f => $"{f.ErrorKey} {f.ErrorDetail}")));
        var hwpx = job.Outputs.Single();
        ValidPackage(hwpx, expectImages);
        await ReadableByH2OrestartAsync(fx, hwpx);
        return hwpx;
    }

    private static int Int(XElement element, string attribute) => (int)element.Attribute(attribute)!;

    [Theory]
    [InlineData("zh-CN", "SimSun", "SimHei")]
    [InlineData("zh-Hant", "MingLiU", "Microsoft JhengHei")]
    [InlineData("ja-JP", "MS Mincho", "Yu Gothic")]
    [InlineData("ko-KR", "Batang", "Malgun Gothic")]
    public void Word_east_asian_theme_fonts_follow_the_document_language(string language, string minor, string major)
    {
        var docx = new DocxBuilder
        {
            Settings = $"<w:themeFontLang w:val=\"en-US\" w:eastAsia=\"{language}\"/>",
            Theme = """
                <a:fontScheme name="Localized">
                  <a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/>
                    <a:font script="Hans" typeface="SimHei"/><a:font script="Hant" typeface="Microsoft JhengHei"/>
                    <a:font script="Jpan" typeface="Yu Gothic"/><a:font script="Hang" typeface="Malgun Gothic"/>
                  </a:majorFont>
                  <a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/>
                    <a:font script="Hans" typeface="SimSun"/><a:font script="Hant" typeface="MingLiU"/>
                    <a:font script="Jpan" typeface="MS Mincho"/><a:font script="Hang" typeface="Batang"/>
                  </a:minorFont>
                </a:fontScheme>
                """,
        };
        docx.Paragraph(P(R("正文", """<w:rFonts w:asciiTheme="minorEastAsia" w:eastAsiaTheme="minorEastAsia"/>""")
            + R("標題", """<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia"/>""")));

        var directory = fx.NewFolder();
        var model = DocxReader.Read(docx.Save(Path.Combine(directory, "theme.docx")), Path.Combine(directory, "media"));
        var paragraph = Assert.IsType<HParagraph>(Assert.Single(Assert.Single(model.Sections).Blocks));
        var text = paragraph.Inlines.OfType<HText>().ToList();
        Assert.Collection(text,
            run => Assert.Equal((minor, minor), (run.Format.Font, run.Format.EastAsianFont)),
            run => Assert.Equal(("Calibri", major), (run.Format.Font, run.Format.EastAsianFont)));
    }

    [Fact]
    public void A_word_center_tab_uses_actual_bold_letter_widths()
    {
        var face = SkiaSharp.SKFontManager.Default.MatchFamily("Cambria", SkiaSharp.SKFontStyle.Bold);
        Assert.SkipUnless(face?.FamilyName == "Cambria", "Cambria is not installed");
        using var font = new SkiaSharp.SKFont(face!, 1600);
        const string title = "AUTHENTICATION SECURITY GUIDANCE";
        var docx = new DocxBuilder().Paragraph(P("""<w:r><w:tab/></w:r>"""
            + R(title, """<w:rFonts w:ascii="Cambria"/><w:b/><w:sz w:val="32"/>"""),
            """<w:tabs><w:tab w:val="center" w:pos="5000"/></w:tabs>"""));
        var dir = fx.NewFolder();
        var model = DocxReader.Read(docx.Save(Path.Combine(dir, "center-tab.docx")), Path.Combine(dir, "media"));
        var path = Path.Combine(dir, "center-tab.hwpx");
        HwpxWriter.Write(model, path, forRhwpPdf: true);

        var tab = Assert.Single(Xml(path).Descendants(Hp + "tab"));
        var expected = (int)(25000 - font.MeasureText(title) / 2);
        Assert.InRange(Int(tab, "width"), expected - 1, expected + 1);
    }

    [Fact]
    public void Word_contents_links_keep_the_entry_style_and_their_destination()
    {
        var docx = new DocxBuilder();
        const string linkRun = """<w:rStyle w:val="Hyperlink"/>""";
        var field = """
            <w:r><w:fldChar w:fldCharType="begin"/></w:r>
            <w:r><w:instrText> TOC \h </w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate"/></w:r>
            """;
        var entry = """<w:hyperlink w:anchor="_Toc1">""" + R("Contents entry", linkRun) + "</w:hyperlink>";
        docx.Paragraph(P(field + entry + """<w:r><w:fldChar w:fldCharType="end"/></w:r>"""));
        var external = docx.Hyperlink("https://example.com/");
        docx.Paragraph(P($"<w:hyperlink r:id=\"{external}\">{R("External link", linkRun)}</w:hyperlink>"));

        var directory = fx.NewFolder();
        var model = DocxReader.Read(docx.Save(Path.Combine(directory, "toc.docx")), Path.Combine(directory, "media"));
        var links = HDocumentWalker.Inlines(Assert.Single(model.Sections).Blocks).OfType<HLink>().ToArray();
        Assert.Equal(2, links.Length);
        Assert.Equal("#_Toc1", links[0].Target);
        var contents = Assert.IsType<HText>(Assert.Single(links[0].Content));
        Assert.Equal("#000000", contents.Format.Color);
        Assert.False(contents.Format.Underline);
        var normal = Assert.IsType<HText>(Assert.Single(links[1].Content));
        Assert.Equal("#0563C1", normal.Format.Color);
        Assert.True(normal.Format.Underline);
    }

    [Fact]
    public async Task Page_setup_headers_footers_and_page_numbers()
    {
        var docx = new DocxBuilder();
        var header = docx.Header(P(R("머리말 제목"), "<w:jc w:val=\"right\"/>"));
        var footer = docx.Footer(P(R("- ") + Field("PAGE", "1") + R(" / ") + Field("NUMPAGES", "3") + R(" -"), "<w:jc w:val=\"center\"/>"));
        docx.Paragraph(P(R("가로 방향 본문")));
        docx.SectionProperties =
            $"<w:headerReference w:type=\"default\" r:id=\"{header}\"/><w:footerReference w:type=\"default\" r:id=\"{footer}\"/>" +
            "<w:pgSz w:w=\"16838\" w:h=\"11906\" w:orient=\"landscape\"/>" +
            "<w:pgMar w:top=\"1985\" w:right=\"1134\" w:bottom=\"1418\" w:left=\"1134\" w:header=\"851\" w:footer=\"851\" w:gutter=\"0\"/><w:titlePg/>";

        var hwpx = await ConvertAsync(docx, "page-setup");
        var section = Xml(hwpx);

        // Landscape: portrait-sized sheet turned by "NARROWLY".
        var pagePr = section.Descendants(Hp + "pagePr").Single();
        Assert.Equal("NARROWLY", (string?)pagePr.Attribute("landscape"));
        Assert.Equal((11906 * 5, 16838 * 5), (Int(pagePr, "width"), Int(pagePr, "height")));
        // Word measures the body from the paper edge; 한글 stacks header area on top of the top margin.
        var margin = pagePr.Element(Hp + "margin")!;
        Assert.Equal(851 * 5, Int(margin, "top"));
        Assert.Equal((1985 - 851) * 5, Int(margin, "header"));
        Assert.Equal(851 * 5, Int(margin, "bottom"));
        Assert.Equal((1418 - 851) * 5, Int(margin, "footer"));
        Assert.Equal(1134 * 5, Int(margin, "left"));

        var headerControl = section.Descendants(Hp + "header").Single();
        Assert.Equal("BOTH", (string?)headerControl.Attribute("applyPageType"));
        Assert.Contains("머리말 제목", headerControl.Value);
        var footerControl = section.Descendants(Hp + "footer").Single();
        var numbers = footerControl.Descendants(Hp + "autoNum").Select(a => (string?)a.Attribute("numType")).ToList();
        Assert.Equal(["PAGE", "TOTAL_PAGE"], numbers);
        Assert.DoesNotContain("3", string.Concat(footerControl.Descendants(Hp + "t").Select(t => t.Value))); // cached field results are dropped

        // "Different first page" without a first-page header hides header and footer on page 1.
        var visibility = section.Descendants(Hp + "visibility").Single();
        Assert.Equal(("1", "1"), ((string?)visibility.Attribute("hideFirstHeader"), (string?)visibility.Attribute("hideFirstFooter")));
        RenderWithRhwp(fx, hwpx, "docx-page-setup");
    }

    [Fact]
    public async Task Columns_change_within_a_section_and_new_paper_starts_a_new_section()
    {
        var docx = new DocxBuilder()
            .Paragraph(P(R("한 단 본문")))
            .Paragraph(P(R("첫 구역 끝"), $"<w:sectPr>{DefaultPage}</w:sectPr>"))
            .Paragraph(P(R("두 단 왼쪽")))
            .Paragraph(P(R("두 단 끝"), $"<w:sectPr><w:type w:val=\"continuous\"/>{DefaultPage}<w:cols w:num=\"2\" w:space=\"720\" w:sep=\"1\"/></w:sectPr>"))
            .Paragraph(P(R("레터 용지")));
        docx.SectionProperties = "<w:type w:val=\"nextPage\"/><w:pgSz w:w=\"12240\" w:h=\"15840\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/>";

        var hwpx = await ConvertAsync(docx, "sections");

        var first = Xml(hwpx);
        var columnChange = ParagraphWith(first, "두 단 왼쪽").Descendants(Hp + "colPr").Single();
        Assert.Equal((2, 3600), (Int(columnChange, "colCount"), Int(columnChange, "sameGap")));
        Assert.NotNull(columnChange.Element(Hp + "colLine"));
        Assert.Equal(1, Int(first.Descendants(Hp + "colPr").First(), "colCount"));

        var second = Xml(hwpx, "Contents/section1.xml");
        var pagePr = second.Descendants(Hp + "pagePr").Single();
        Assert.Equal(("WIDELY", 61200, 79200), ((string)pagePr.Attribute("landscape")!, Int(pagePr, "width"), Int(pagePr, "height")));
        Assert.Contains("레터 용지", second.Value);
        Assert.Equal(2, RenderWithRhwp(fx, hwpx, "docx-sections") ?? 2);
    }

    [Fact]
    public async Task Contextual_spacing_removes_the_gap_between_paragraphs_of_the_same_style()
    {
        var docx = new DocxBuilder();
        docx.Styles = docx.Styles.Replace("</w:docDefaults>",
            "</w:docDefaults><w:style w:type=\"paragraph\" w:styleId=\"ListBullet\"><w:name w:val=\"List Bullet\"/><w:basedOn w:val=\"Normal\"/>" +
            "<w:pPr><w:spacing w:before=\"100\" w:after=\"200\"/><w:contextualSpacing/></w:pPr></w:style>");
        docx.Paragraph(P(R("목록 하나"), "<w:pStyle w:val=\"ListBullet\"/>"))
            .Paragraph(P(R("목록 둘"), "<w:pStyle w:val=\"ListBullet\"/>"))
            .Paragraph(P(R("목록 셋"), "<w:pStyle w:val=\"ListBullet\"/>"))
            .Paragraph(P(R("일반 문단")));

        var hwpx = await ConvertAsync(docx, "contextual-spacing");
        var section = Xml(hwpx);
        var head = Xml(hwpx, "Contents/header.xml");
        XElement ParaPr(XElement paragraph) => head.Descendants(Hh + "paraPr").Single(p => (string?)p.Attribute("id") == (string?)paragraph.Attribute("paraPrIDRef"));
        (int Before, int After) Gaps(string text) =>
            ((int)ParaPr(ParagraphWith(section, text)).Descendants(Hp + "case").Single().Descendants(Hc + "prev").Single().Attribute("value")!,
             (int)ParaPr(ParagraphWith(section, text)).Descendants(Hp + "case").Single().Descendants(Hc + "next").Single().Attribute("value")!);

        // Between two items of the same style there is no gap; the first keeps its space above, the last the space below.
        Assert.Equal((100 * 5, 0), Gaps("목록 하나"));
        Assert.Equal((0, 0), Gaps("목록 둘"));
        Assert.Equal((0, 200 * 5), Gaps("목록 셋"));
    }

    [Fact]
    public async Task Adjacent_word_paragraphs_use_the_larger_gap()
    {
        var docx = new DocxBuilder()
            .Paragraph(P(R("First line"), """<w:spacing w:after="480"/>"""))
            .Paragraph(P(R("Second line"), """<w:spacing w:before="480" w:after="240"/>"""))
            .Paragraph(P(R("Third line"), """<w:spacing w:before="480"/>"""));
        var hwpx = await ConvertAsync(docx, "collapsed-paragraph-spacing");
        var section = Xml(hwpx);
        var header = Xml(hwpx, "Contents/header.xml");
        (int Before, int After) Spacing(string text)
        {
            var id = (string?)ParagraphWith(section, text).Attribute("paraPrIDRef");
            var paraPr = header.Descendants(Hh + "paraPr").Single(p => (string?)p.Attribute("id") == id);
            var margin = paraPr.Descendants(Hp + "case").Single().Descendants(Hh + "margin").Single();
            return (Int(margin.Element(Hc + "prev")!, "value"), Int(margin.Element(Hc + "next")!, "value"));
        }

        Assert.Equal(2400, Spacing("First line").After + Spacing("Second line").Before);
        Assert.Equal(2400, Spacing("Second line").After + Spacing("Third line").Before);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task A_word_run_page_break_does_not_insert_an_empty_line_at_the_top(bool standalone)
    {
        const string pageBreak = """<w:r><w:br w:type="page"/></w:r>""";
        var first = standalone ? P(R("Previous page")) + P(pageBreak) : P(R("Previous page") + pageBreak);
        var docx = new DocxBuilder().Paragraph(first)
            .Paragraph(P(R("New chapter"), """<w:spacing w:before="480"/>"""));
        var hwpx = await ConvertAsync(docx, "page-break-gap");
        var section = Xml(hwpx);
        var chapter = ParagraphWith(section, "New chapter");
        Assert.Equal("1", (string?)chapter.Attribute("pageBreak"));
        Assert.Single(section.Descendants(Hp + "p"), p => (string?)p.Attribute("pageBreak") == "1");
    }

    [Fact]
    public async Task Paragraph_and_character_formatting_lists_and_document_grid()
    {
        var docx = new DocxBuilder
        {
            Numbering = """
                <w:abstractNum w:abstractNumId="1">
                  <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="400" w:hanging="400"/></w:pPr></w:lvl>
                  <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="ganada"/><w:lvlText w:val="%2)"/><w:pPr><w:ind w:left="800" w:hanging="400"/></w:pPr></w:lvl>
                </w:abstractNum>
                <w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr></w:lvl></w:abstractNum>
                <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
                <w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>
                """,
            SectionProperties = DefaultPage + "<w:docGrid w:type=\"lines\" w:linePitch=\"360\"/>",
        };
        docx.Paragraph(P(R("가운데 제목"), "<w:pStyle w:val=\"Heading1\"/><w:jc w:val=\"center\"/>"))
            .Paragraph(P(R("굵은 빨강", "<w:b/><w:color w:val=\"FF0000\"/><w:sz w:val=\"28\"/><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\" w:eastAsia=\"굴림\"/>") +
                         R(" 형광펜", "<w:highlight w:val=\"yellow\"/>") + R(" 무늬", "<w:shd w:val=\"pct25\" w:color=\"auto\" w:fill=\"auto\"/>"),
                "<w:ind w:left=\"720\" w:hanging=\"360\"/><w:spacing w:before=\"240\" w:after=\"120\"/>"))
            .Paragraph(P(R("첫 항목"), "<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>"))
            .Paragraph(P(R("하위 항목"), "<w:numPr><w:ilvl w:val=\"1\"/><w:numId w:val=\"1\"/></w:numPr>"))
            .Paragraph(P(R("글머리 항목"), "<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"2\"/></w:numPr>"))
            .Paragraph(P(R("2") + R("016-06-15")));

        var hwpx = await ConvertAsync(docx, "formatting");
        var section = Xml(hwpx);
        var head = Xml(hwpx, "Contents/header.xml");
        XElement ParaPr(XElement paragraph) => head.Descendants(Hh + "paraPr").Single(p => (string?)p.Attribute("id") == (string?)paragraph.Attribute("paraPrIDRef"));
        XElement CharPr(XElement run) => head.Descendants(Hh + "charPr").Single(c => (string?)c.Attribute("id") == (string?)run.Attribute("charPrIDRef"));
        int Margin(XElement paraPr, string branch, string name) =>
            (int)paraPr.Descendants(Hp + branch).Single().Descendants(Hc + name).Single().Attribute("value")!;

        var title = ParagraphWith(section, "가운데 제목");
        var headingStyle = head.Descendants(Hh + "style").Single(s => (string?)s.Attribute("engName") == "Heading 1");
        Assert.Equal((string?)headingStyle.Attribute("id"), (string?)title.Attribute("styleIDRef"));
        Assert.Equal("CENTER", (string?)ParaPr(title).Element(Hh + "align")!.Attribute("horizontal"));

        var red = CharPr(RunWith(section, "굵은 빨강"));
        Assert.Equal(("1400", "#FF0000"), ((string)red.Attribute("height")!, (string)red.Attribute("textColor")!));
        Assert.NotNull(red.Element(Hh + "bold"));
        string Font(string language) =>
            (string)head.Descendants(Hh + "fontface").Single(f => (string?)f.Attribute("lang") == language)
                .Elements(Hh + "font").Single(f => (string?)f.Attribute("id") == (string?)red.Element(Hh + "fontRef")!.Attribute(language.ToLowerInvariant())).Attribute("face")!;
        Assert.Equal(("Arial", "굴림"), (Font("LATIN"), Font("HANGUL")));
        Assert.Equal("#FFFF00", (string?)CharPr(RunWith(section, "형광펜")).Attribute("shadeColor"));
        Assert.Equal("#BFBFBF", (string?)CharPr(RunWith(section, "무늬")).Attribute("shadeColor")); // 25% black on white

        // Hanging indent (Word: text at 0.5", first line at 0.25"). 한글 keeps a hanging paragraph's first line at
        // "left" and indents the rest, so left is the first-line position. hp:default repeats the values doubled.
        var indented = ParaPr(ParagraphWith(section, "굵은 빨강"));
        Assert.Equal((1800, -1800, 1200, 600), (Margin(indented, "case", "left"), Margin(indented, "case", "intent"), Margin(indented, "case", "prev"), Margin(indented, "case", "next")));
        Assert.Equal((3600, -3600), (Margin(indented, "default", "left"), Margin(indented, "default", "intent")));

        // Word's 18 pt document grid: 10 pt lines become 18 pt, 14 pt lines need two grid lines.
        var gridLine = ParaPr(ParagraphWith(section, "첫 항목")).Descendants(Hp + "case").Single().Element(Hh + "lineSpacing")!;
        Assert.Equal(("AT_LEAST", 1800), ((string)gridLine.Attribute("type")!, Int(gridLine, "value")));
        Assert.Equal(3600, Int(indented.Descendants(Hp + "case").Single().Element(Hh + "lineSpacing")!, "value"));

        // Lists: Word levels map to 한글 numbering levels; bullets use a symbol instead of a counter.
        var item = ParaPr(ParagraphWith(section, "첫 항목")).Element(Hh + "heading")!;
        var subItem = ParaPr(ParagraphWith(section, "하위 항목")).Element(Hh + "heading")!;
        Assert.Equal(("NUMBER", 0, 1), ((string)item.Attribute("type")!, Int(item, "level"), Int(subItem, "level")));
        var numbering = head.Descendants(Hh + "numbering").Single(n => (string?)n.Attribute("id") == (string?)item.Attribute("idRef"));
        var levels = numbering.Elements(Hh + "paraHead").Take(2).Select(h => ((string?)h.Attribute("numFormat"), h.Value)).ToList();
        Assert.Equal([("DIGIT", "^1."), ("HANGUL_SYLLABLE", "^2)")], levels);
        var bullet = ParaPr(ParagraphWith(section, "글머리 항목")).Element(Hh + "heading")!;
        Assert.Equal("•", head.Descendants(Hh + "numbering").Single(n => (string?)n.Attribute("id") == (string?)bullet.Attribute("idRef")).Elements(Hh + "paraHead").First().Value);

        // Runs split by Word but formatted alike become one text run.
        Assert.Contains(section.Descendants(Hp + "t"), t => t.Value == "2016-06-15");
        RenderWithRhwp(fx, hwpx, "docx-formatting");
    }

    [Fact]
    public async Task Bordered_word_table_rows_reserve_their_share_of_each_horizontal_border()
    {
        static string Table(string style, int eighths, string tableProperties = "", string rowProperties = "")
        {
            var cell = $"<w:tc>{P(R("Row text"))}</w:tc>";
            var row = $"<w:tr><w:trPr>{rowProperties}</w:trPr>{cell}</w:tr>";
            return $"<w:tbl><w:tblPr>{tableProperties}{HorizontalBorders(style, eighths)}</w:tblPr>" +
                   $"<w:tblGrid><w:gridCol w:w=\"3000\"/></w:tblGrid>{row}{row}</w:tbl>";
        }

        const string fixedHeight = """<w:trHeight w:val="20" w:hRule="exact"/>""";
        const string minimumHeight = """<w:trHeight w:val="369"/>""";
        var docx = new DocxBuilder()
            .Paragraph(Table("nil", 8))
            .Paragraph(Table("single", 4))
            .Paragraph(Table("single", 8))
            .Paragraph(Table("single", 8, """<w:tblCellSpacing w:w="20" w:type="dxa"/>"""))
            .Paragraph(Table("nil", 8, rowProperties: fixedHeight))
            .Paragraph(Table("single", 8, rowProperties: fixedHeight))
            .Paragraph(Table("single", 4, rowProperties: minimumHeight))
            .Paragraph(P(""));
        var hwpx = await ConvertAsync(docx, "table-border-heights");
        var rowHeights = Xml(hwpx).Descendants(Hp + "tbl")
            .Select(table => table.Descendants(Hp + "cellSz").Select(size => Int(size, "height")).ToArray())
            .ToArray();

        Assert.Equal(rowHeights[0].Select(height => height + 50), rowHeights[1]);
        Assert.Equal(rowHeights[0].Select(height => height + 100), rowHeights[2]);
        Assert.Equal(rowHeights[0], rowHeights[3]);
        Assert.Equal(rowHeights[4], rowHeights[5]);
        Assert.Equal([1895, 1895], rowHeights[6]);
    }

    [Fact]
    public async Task Built_in_word_pdf_paginates_bordered_rows_with_a_repeated_header()
    {
        var rhwp = fx.Converters.Single(c => c.Id == "rhwp");
        Assert.SkipUnless(fx.Catalog.StatusOf(rhwp).IsAvailable, "rhwp is not available");
        var builtIn = new ConverterCatalog(fx.Converters.Where(c => c.Id is "hwpx-writer" or "rhwp"));
        var docx = new DocxBuilder
        {
            SectionProperties = DefaultPage.Replace("16838", "17038", StringComparison.Ordinal),
        };
        docx.Body.Append($"<w:tbl><w:tblPr>{HorizontalBorders("single", 4)}");
        docx.Body.Append("""<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar>""");
        docx.Body.Append("""</w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>""");
        for (var i = 0; i < 40; i++)
        {
            var text = i == 0 ? "TableHeader" : $"DataRow{i:00}";
            var header = i == 0 ? "<w:tblHeader/>" : "";
            var run = R(text, """<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/>""");
            var paragraph = P(run, """<w:spacing w:after="0" w:line="412" w:lineRule="exact"/>""");
            docx.Body.Append($"<w:tr><w:trPr>{header}</w:trPr><w:tc>{paragraph}</w:tc></w:tr>");
        }
        docx.Body.Append("</w:tbl>").Append(P(""));

        var input = docx.Save(Path.Combine(fx.NewFolder(), "paginated-table.docx"));
        var job = await fx.ConvertAsync([input], new Preset { TargetFormat = "pdf" }, builtIn);
        Assert.True(job.State == JobState.Completed, string.Join("; ", job.Files.Select(f => $"{f.ErrorKey} {f.ErrorDetail}")));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(job.Outputs.Single());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("DataRow31", pdf.GetPage(1).Text);
        Assert.DoesNotContain("DataRow32", pdf.GetPage(1).Text);
        Assert.Contains("DataRow32", pdf.GetPage(2).Text);
        Assert.Contains("TableHeader", pdf.GetPage(2).Text);
    }

    [Fact]
    public async Task Inline_word_table_does_not_add_a_body_paragraph_gap()
    {
        var docx = new DocxBuilder()
            .Paragraph(P(R("Before table")))
            .Paragraph($"<w:tbl><w:tr><w:tc>{P(R("One cell"))}</w:tc></w:tr></w:tbl>")
            .Paragraph(P(R("After table")));
        var hwpx = await ConvertAsync(docx, "inline-table-gap");
        var table = Assert.Single(Xml(hwpx).Descendants(Hp + "tbl"));
        var run = table.Parent!;
        var header = Xml(hwpx, "Contents/header.xml");
        var charId = (string?)run.Attribute("charPrIDRef");
        var charPr = header.Descendants(Hh + "charPr").Single(p => (string?)p.Attribute("id") == charId);
        Assert.Equal(0, Int(table.Element(Hp + "outMargin")!, "bottom"));
        Assert.Equal(100, Int(charPr, "height"));
    }

    [Fact]
    public async Task Numeric_ranges_that_fit_the_cell_do_not_reserve_a_second_line()
    {
        const string face = "Times New Roman";
        Assert.SkipUnless(Filee.Engines.Hwp.RhwpFonts.Find(face) is { Usable: true }, $"{face} is not installed");
        const string runProperties = """<w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="24"/>""";
        const string tableProperties = """
            <w:tblPr><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/>
              <w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr>
            <w:tblGrid><w:gridCol w:w="1418"/></w:tblGrid>
            """;
        static string Row(string value) => $"<w:tr><w:tc>{P(R(value, runProperties))}</w:tc></w:tr>";
        var docx = new DocxBuilder().Paragraph(
            $"<w:tbl>{tableProperties}{Row("0–1,479")}{Row("1,480–2,959")}{Row("1,480–2,959–3,979")}</w:tbl>");
        var hwpx = await ConvertAsync(docx, "numeric-row-height");
        var heights = Xml(hwpx).Descendants(Hp + "cellSz").Select(cell => Int(cell, "height")).ToArray();

        // The 60 pt range fits in 60.1 pt of content width, while the longer range needs another line.
        Assert.Equal(heights[0], heights[1]);
        Assert.True(heights[2] > heights[1]);
    }

    private static string HorizontalBorders(string style, int eighths) =>
        "<w:tblBorders>" + string.Concat(new[] { "top", "insideH", "bottom" }
            .Select(side => $"<w:{side} w:val=\"{style}\" w:sz=\"{eighths}\"/>")) + "</w:tblBorders>";

    [Fact]
    public async Task Tables_keep_merged_cells_borders_and_shading()
    {
        var cell = (string content, string tcPr) => $"<w:tc><w:tcPr>{tcPr}</w:tcPr>{P(R(content))}</w:tc>";
        var docx = new DocxBuilder().Paragraph(
            "<w:tbl><w:tblPr><w:tblStyle w:val=\"TableGrid\"/><w:jc w:val=\"center\"/></w:tblPr>" +
            "<w:tblGrid><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"3000\"/><w:gridCol w:w=\"2000\"/></w:tblGrid>" +
            "<w:tr><w:trPr><w:tblHeader/></w:trPr>" + cell("가로 병합", "<w:gridSpan w:val=\"2\"/><w:shd w:val=\"pct10\" w:color=\"auto\" w:fill=\"auto\"/>") +
                cell("세로 병합", "<w:vMerge w:val=\"restart\"/><w:vAlign w:val=\"center\"/>") + "</w:tr>" +
            "<w:tr>" + cell("A", "") + cell("B", "") + cell("", "<w:vMerge/>") + "</w:tr>" +
            "<w:tr>" + cell("C", "<w:tcBorders><w:bottom w:val=\"double\" w:sz=\"12\" w:color=\"FF0000\"/></w:tcBorders>") + cell("D", "<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"FFF2CC\"/>") + cell("E", "") + "</w:tr>" +
            "</w:tbl>" + P(""));

        var hwpx = await ConvertAsync(docx, "table");
        var section = Xml(hwpx);
        var table = section.Descendants(Hp + "tbl").Single();
        Assert.Equal((3, 3, "1"), (Int(table, "rowCnt"), Int(table, "colCnt"), (string?)table.Attribute("repeatHeader")));
        Assert.Equal("CENTER", (string?)table.Element(Hp + "pos")!.Attribute("horzAlign"));

        var cells = table.Descendants(Hp + "tc").Select(tc => new
        {
            Text = tc.Element(Hp + "subList")!.Value,
            Row = Int(tc.Element(Hp + "cellAddr")!, "rowAddr"),
            Column = Int(tc.Element(Hp + "cellAddr")!, "colAddr"),
            RowSpan = Int(tc.Element(Hp + "cellSpan")!, "rowSpan"),
            ColSpan = Int(tc.Element(Hp + "cellSpan")!, "colSpan"),
            Width = Int(tc.Element(Hp + "cellSz")!, "width"),
            BorderFill = HeadItem(hwpx, "borderFill", (string)tc.Attribute("borderFillIDRef")!),
            VerticalAlign = (string?)tc.Element(Hp + "subList")!.Attribute("vertAlign"),
        }).ToList();
        // Every grid position is covered exactly once.
        var covered = cells.SelectMany(c => Enumerable.Range(c.Row, c.RowSpan).SelectMany(r => Enumerable.Range(c.Column, c.ColSpan).Select(col => (r, col)))).ToList();
        Assert.Equal(9, covered.Distinct().Count());
        Assert.Equal(9, covered.Count);

        var wide = cells.Single(c => c.Text == "가로 병합");
        Assert.Equal((2, 25000), (wide.ColSpan, wide.Width));
        Assert.Equal("#E6E6E6", (string?)wide.BorderFill.Descendants(Hc + "winBrush").Single().Attribute("faceColor"));
        var tall = cells.Single(c => c.Text == "세로 병합");
        Assert.Equal((2, "CENTER"), (tall.RowSpan, tall.VerticalAlign));
        var bottom = cells.Single(c => c.Text == "C").BorderFill.Element(Hh + "bottomBorder")!;
        Assert.Equal(("DOUBLE_SLIM", "#FF0000"), ((string)bottom.Attribute("type")!, (string)bottom.Attribute("color")!));
        Assert.Equal("SOLID", (string?)cells.Single(c => c.Text == "A").BorderFill.Element(Hh + "leftBorder")!.Attribute("type"));
        Assert.Equal("#FFF2CC", (string?)cells.Single(c => c.Text == "D").BorderFill.Descendants(Hc + "winBrush").Single().Attribute("faceColor"));
        RenderWithRhwp(fx, hwpx, "docx-table");
    }

    [Fact]
    public async Task A_right_tab_accounts_for_the_inline_image_before_it()
    {
        var docx = new DocxBuilder();
        using var image = new MagickImage(MagickColors.Black, 10, 10);
        var picture = docx.Image(image.ToByteArray(MagickFormat.Png));
        const string tabStop = """<w:tabs><w:tab w:val="right" w:pos="7200"/></w:tabs>""";
        const string tab = "<w:r><w:tab/></w:r>";
        docx.Paragraph(P(R("Signed") + tab + R("2026"), tabStop));
        docx.Paragraph(P(R("Signed") + InlinePicture(picture, 254000, 127000) + tab + R("2026"), tabStop));

        var hwpx = await ConvertAsync(docx, "tab-after-inline-image", expectImages: 1);
        var tabs = Xml(hwpx).Descendants(Hp + "tab").ToArray();
        Assert.Equal(2, tabs.Length);
        Assert.Equal(2000, Int(tabs[0], "width") - Int(tabs[1], "width"));
        Assert.All(tabs, t => Assert.Equal(2, Int(t, "type")));

        var folder = fx.NewFolder();
        var source = docx.Save(Path.Combine(folder, "source.docx"));
        var renderInput = Path.Combine(folder, "pdf-input.hwpx");
        HwpxWriter.Write(DocxReader.Read(source, Path.Combine(folder, "media")), renderInput, forRhwpPdf: true);
        Assert.Equal([2, 1], Xml(renderInput).Descendants(Hp + "tab").Select(t => Int(t, "type")));
    }

    [Fact]
    public async Task Text_boxes_shapes_and_floating_pictures()
    {
        var docx = new DocxBuilder();
        using var image = new MagickImage(MagickColors.MediumPurple, 120, 60);
        var picture = docx.Image(image.ToByteArray(MagickFormat.Png));
        const string Column = "<wp:positionH relativeFrom=\"column\"><wp:posOffset>0</wp:posOffset></wp:positionH>";
        const string Paragraph = "<wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>0</wp:posOffset></wp:positionV>";
        string Shape(string geometry, string fill, string line, long cx, long cy, string text = "") =>
            $"<wps:wsp><wps:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"{cx}\" cy=\"{cy}\"/></a:xfrm><a:prstGeom prst=\"{geometry}\"/>{fill}{line}</wps:spPr>" +
            (text.Length > 0 ? $"<wps:txbx><w:txbxContent>{P(R(text))}</w:txbxContent></wps:txbx>" : "") + "<wps:bodyPr anchor=\"ctr\"/></wps:wsp>";

        docx.Paragraph(P(Anchor(PictureUri, Picture(picture, 914400, 457200), 914400, 457200,
                "<wp:positionH relativeFrom=\"page\"><wp:posOffset>914400</wp:posOffset></wp:positionH>", Paragraph) + R("그림 옆에 흐르는 글")))
            .Paragraph(P(InlinePicture(picture, 457200, 457200)))
            .Paragraph(P(Anchor(PictureUri, Picture(picture, 457200, 457200), 457200, 457200,
                "<wp:positionH relativeFrom=\"margin\"><wp:posOffset>-635000</wp:posOffset></wp:positionH>",
                "<wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>-127000</wp:posOffset></wp:positionV>") + R("여백으로 나간 그림")))
            .Paragraph(P(Anchor(ShapeUri,
                Shape("rect", "<a:solidFill><a:srgbClr val=\"FFF2CC\"/></a:solidFill>", "<a:ln w=\"12700\"><a:solidFill><a:srgbClr val=\"C00000\"/></a:solidFill></a:ln>", 1828800, 914400, "글상자 안의 글"),
                1828800, 914400, "<wp:positionH relativeFrom=\"margin\"><wp:align>center</wp:align></wp:positionH>",
                "<wp:positionV relativeFrom=\"page\"><wp:posOffset>3600000</wp:posOffset></wp:positionV>", "<wp:wrapTopAndBottom/>")))
            .Paragraph(P(Anchor(ShapeUri, Shape("line", "", "<a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"4F81BD\"/></a:solidFill></a:ln>", 3000000, 0),
                3000000, 0, Column, Paragraph, "<wp:wrapNone/>")))
            .Paragraph(P(Anchor(ShapeUri,
                Shape("rect", "<a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"99CCFF\"/></a:gs><a:gs pos=\"100000\"><a:srgbClr val=\"FFFFFF\"/></a:gs></a:gsLst><a:lin ang=\"0\"/></a:gradFill>", "<a:ln><a:noFill/></a:ln>", 2000000, 180000),
                2000000, 180000, Column, Paragraph, "<wp:wrapNone/>")));

        var hwpx = await ConvertAsync(docx, "objects", expectImages: 1); // one picture used twice is stored once
        var section = Xml(hwpx);

        var pictures = section.Descendants(Hp + "pic").ToList();
        var floating = pictures.Single(p => (string?)p.Element(Hp + "pos")!.Attribute("treatAsChar") == "0" && Int(p.Element(Hp + "pos")!, "horzOffset") == 7200);
        Assert.Equal(("SQUARE", "PAPER", 7200), ((string)floating.Attribute("textWrap")!, (string)floating.Element(Hp + "pos")!.Attribute("horzRelTo")!, Int(floating.Element(Hp + "pos")!, "horzOffset")));
        Assert.Single(pictures, p => (string?)p.Element(Hp + "pos")!.Attribute("treatAsChar") == "1");
        // Negative offsets are not allowed in 한글 files: into the margin → measured from the paper edge instead.
        var outside = pictures.Single(p => (string?)p.Element(Hp + "pos")!.Attribute("horzRelTo") == "PAPER" && Int(p.Element(Hp + "pos")!, "horzOffset") != 7200).Element(Hp + "pos")!;
        Assert.Equal((1440 * 5 - 5000, 0), (Int(outside, "horzOffset"), Int(outside, "vertOffset")));
        Assert.DoesNotContain(section.Descendants(Hp + "pos"), p => Int(p, "horzOffset") < 0 || Int(p, "vertOffset") < 0);

        var textBox = section.Descendants(Hp + "rect").Single(r => r.Element(Hp + "drawText") is not null);
        Assert.Contains("글상자 안의 글", textBox.Element(Hp + "drawText")!.Value);
        Assert.Equal("CENTER", (string?)textBox.Descendants(Hp + "subList").First().Attribute("vertAlign"));
        Assert.Equal("#FFF2CC", (string?)textBox.Descendants(Hc + "winBrush").Single().Attribute("faceColor"));
        Assert.Equal(("#C00000", 100), ((string)textBox.Element(Hp + "lineShape")!.Attribute("color")!, Int(textBox.Element(Hp + "lineShape")!, "width")));
        var position = textBox.Element(Hp + "pos")!;
        Assert.Equal(("PAGE", "CENTER", "PAPER", 28346), ((string)position.Attribute("horzRelTo")!, (string)position.Attribute("horzAlign")!, (string)position.Attribute("vertRelTo")!, Int(position, "vertOffset")));
        Assert.Equal("TOP_AND_BOTTOM", (string?)textBox.Attribute("textWrap"));

        var line = section.Descendants(Hp + "line").Single();
        Assert.Equal(23622, Int(line.Element(Hc + "endPt")!, "x"));
        Assert.Equal("#4F81BD", (string?)line.Element(Hp + "lineShape")!.Attribute("color"));
        var gradient = section.Descendants(Hp + "rect").Single(r => r.Element(Hp + "drawText") is null).Descendants(Hc + "gradation").Single();
        Assert.Equal(["#99CCFF", "#FFFFFF"], gradient.Elements(Hc + "color").Select(c => (string)c.Attribute("value")!));
        RenderWithRhwp(fx, hwpx, "docx-objects");
    }

    [Fact]
    public async Task Notes_links_bookmarks_and_page_breaks()
    {
        var docx = new DocxBuilder();
        docx.Footnotes("<w:footnote w:id=\"1\"><w:p><w:r><w:footnoteRef/></w:r>" + R(" 각주 설명") + "</w:p></w:footnote>");
        var url = docx.Hyperlink("https://example.com/a?b=1");
        docx.Paragraph(P("<w:bookmarkStart w:id=\"0\" w:name=\"intro\"/>" + R("소개") + "<w:bookmarkEnd w:id=\"0\"/>", "<w:pStyle w:val=\"Heading1\"/>"))
            .Paragraph(P(R("본문") + "<w:r><w:footnoteReference w:id=\"1\"/></w:r>" +
                         $"<w:hyperlink r:id=\"{url}\">" + R("웹 링크", "<w:rStyle w:val=\"Hyperlink\"/>") + "</w:hyperlink>" +
                         "<w:hyperlink w:anchor=\"intro\">" + R("소개로 이동") + "</w:hyperlink>"))
            .Paragraph(P(R("앞 쪽") + "<w:r><w:br w:type=\"page\"/></w:r>" + R("다음 쪽")));

        var hwpx = await ConvertAsync(docx, "notes-links");
        var section = Xml(hwpx);

        // The note number is an autoNum control at the start of the note's first paragraph.
        var note = section.Descendants(Hp + "footNote").Single();
        var firstRun = note.Descendants(Hp + "p").First().Elements(Hp + "run").First();
        Assert.Equal("FOOTNOTE", (string?)firstRun.Descendants(Hp + "autoNum").Single().Attribute("numType"));
        Assert.Contains("각주 설명", note.Value);

        string Parameter(XElement field, string name) => field.Descendants().Single(p => (string?)p.Attribute("name") == name).Value;
        var links = section.Descendants(Hp + "fieldBegin").ToList();
        var web = links.Single(l => Parameter(l, "Category") == "HWPHYPERLINK_TYPE_URL");
        Assert.Equal("https\\://example.com/a\\?b=1;1;0;0;", Parameter(web, "Command"));
        Assert.Equal("https://example.com/a?b=1", Parameter(web, "Path"));
        var bookmarkLink = links.Single(l => Parameter(l, "Category") == "HWPHYPERLINK_TYPE_HWP");
        Assert.Equal("?intro;0;0;0;", Parameter(bookmarkLink, "Command"));
        Assert.Equal("intro", (string?)section.Descendants(Hp + "bookmark").Single().Attribute("name"));
        Assert.Equal(2, section.Descendants(Hp + "fieldEnd").Count());

        Assert.Equal("1", (string?)ParagraphWith(section, "다음 쪽").Attribute("pageBreak"));
        Assert.Equal("0", (string?)ParagraphWith(section, "앞 쪽").Attribute("pageBreak"));
        Assert.Equal(2, RenderWithRhwp(fx, hwpx, "docx-notes-links") ?? 2);
    }
}
