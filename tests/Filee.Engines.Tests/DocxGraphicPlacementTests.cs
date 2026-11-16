using System.IO.Compression;
using Filee.Core.Conversion;
using Filee.Core.Presets;
using ImageMagick;
using UglyToad.PdfPig;
using static Filee.Engines.Tests.DocxBuilder;
using static Filee.Engines.Tests.HwpxAssert;

namespace Filee.Engines.Tests;

public class DocxGraphicPlacementTests(EngineFixture fx) : IClassFixture<EngineFixture>
{
    private async Task<string> ConvertAsync(DocxBuilder docx, string name)
    {
        var source = docx.Save(Path.Combine(fx.NewFolder(), name + ".docx"));
        var job = await fx.ConvertAsync([source], new Preset { TargetFormat = "hwpx" });
        Assert.True(job.State == JobState.Completed, string.Join("; ", job.Files.Select(f => $"{f.ErrorKey} {f.ErrorDetail}")));
        return job.Outputs.Single();
    }

    [Fact]
    public async Task Page_positioned_frame_keeps_picture_and_caption_together()
    {
        var docx = new DocxBuilder().Paragraph(P(R("Body")));
        using var image = new MagickImage(MagickColors.Black, 50, 50);
        var picture = docx.Image(image.ToByteArray(MagickFormat.Png));
        docx.Paragraph(P(InlinePicture(picture, 925000, 861931) + "<w:r><w:br/></w:r>" + R("Verification code"),
            """
            <w:framePr w:w="1680" w:h="1961" w:hRule="exact" w:wrap="none" w:vAnchor="page" w:hAnchor="page" w:x="9840" w:y="14594"/><w:jc w:val="center"/>
            """));
        var hwpx = await ConvertAsync(docx, "page-frame");
        ValidPackage(hwpx, expectImages: 1);
        var frame = Xml(hwpx).Descendants(Hp + "rect").Single();
        var pos = frame.Element(Hp + "pos")!;
        var size = frame.Element(Hp + "sz")!;
        Assert.Equal(("PAPER", "49200"), ((string?)pos.Attribute("horzRelTo"), (string?)pos.Attribute("horzOffset")));
        Assert.Equal(("PAPER", "72970"), ((string?)pos.Attribute("vertRelTo"), (string?)pos.Attribute("vertOffset")));
        Assert.Equal((8400, 9805), ((int)size.Attribute("width")!, (int)size.Attribute("height")!));
        Assert.Single(frame.Descendants(Hp + "pic"));
        Assert.Contains("Verification code", frame.Value);
        Assert.Equal(1, RenderWithRhwp(fx, hwpx, "docx-page-frame") ?? 1);
    }

    [Fact]
    public async Task Floating_word_table_keeps_its_alignment_and_text_clearance()
    {
        var docx = new DocxBuilder()
            .Paragraph(P(R("Above the table")))
            .Paragraph("""
                <w:tbl><w:tblPr>
                  <w:tblpPr w:horzAnchor="margin" w:tblpXSpec="center" w:vertAnchor="text" w:tblpY="177" w:leftFromText="180" w:rightFromText="220"/>
                  <w:tblOverlap w:val="never"/>
                </w:tblPr><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>
                <w:tr><w:tc><w:p><w:r><w:t>Floating table</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
                """)
            .Paragraph(P(R("Below the table")));
        var hwpx = await ConvertAsync(docx, "floating-table");
        var table = Assert.Single(Xml(hwpx).Descendants(Hp + "tbl"));
        Assert.Equal("SQUARE", (string?)table.Attribute("textWrap"));
        var pos = table.Element(Hp + "pos")!;
        Assert.Equal(("PAGE", "CENTER"), ((string?)pos.Attribute("horzRelTo"), (string?)pos.Attribute("horzAlign")));
        Assert.Equal(("PARA", "880"), ((string?)pos.Attribute("vertRelTo"), (string?)pos.Attribute("vertOffset")));
        Assert.Equal("0", (string?)pos.Attribute("allowOverlap"));
        var margin = table.Element(Hp + "outMargin")!;
        Assert.Equal((900, 1100, 0), ((int)margin.Attribute("left")!, (int)margin.Attribute("right")!, (int)margin.Attribute("bottom")!));
    }

    [Fact]
    public async Task A_table_positioned_on_the_paper_can_extend_into_the_body_margins()
    {
        var docx = new DocxBuilder().Paragraph(P(R("Before the table"))).Paragraph("""
            <w:tbl>
              <w:tblPr><w:tblpPr w:horzAnchor="page" w:tblpX="201" w:vertAnchor="text" w:tblpY="61"/>
                <w:tblW w:w="10700" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>
              <w:tblGrid><w:gridCol w:w="2700"/><w:gridCol w:w="8000"/></w:tblGrid>
              <w:tr><w:tc><w:p><w:r><w:t>First cell</w:t></w:r></w:p></w:tc>
                    <w:tc><w:p><w:r><w:t>Second cell</w:t></w:r></w:p></w:tc></w:tr>
            </w:tbl>
            """);
        var hwpx = await ConvertAsync(docx, "page-wide-table");
        var table = Assert.Single(Xml(hwpx).Descendants(Hp + "tbl"));
        Assert.Equal(53500, (int)table.Element(Hp + "sz")!.Attribute("width")!);
        Assert.Equal([13500, 40000], table.Descendants(Hp + "cellSz").Select(cell => (int)cell.Attribute("width")!));
    }

    [Fact]
    public async Task Vml_watermark_has_a_rotated_translucent_picture_and_hidden_shapes_are_omitted()
    {
        var docx = new DocxBuilder();
        var header = docx.Header(P("""
            <w:r><w:pict>
              <v:shape style="position:absolute;width:100pt;height:50pt;visibility:hidden">
                <v:textpath on="t" string="hidden shape"/>
              </v:shape>
              <v:shape style="position:absolute;width:240pt;height:36pt;rotation:315;z-index:-1;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" fillcolor="#a5a5a5" stroked="f">
                <v:fill opacity="29491f"/>
                <v:textpath style="font-family:Arial;font-size:32pt;font-weight:bold" string="CONFIDENTIAL"/>
              </v:shape>
            </w:pict></w:r>
            """));
        docx.SectionProperties = $"<w:headerReference w:type=\"default\" r:id=\"{header}\"/>" + DefaultPage;
        docx.Paragraph(P(R("Body")));

        var hwpx = await ConvertAsync(docx, "vml-watermark");
        ValidPackage(hwpx, expectImages: 1);
        var headerControl = Xml(hwpx).Descendants(Hp + "header").Single();
        Assert.DoesNotContain("hidden shape", headerControl.Value);
        var mark = headerControl.Descendants(Hp + "pic").Single();
        Assert.Equal("BEHIND_TEXT", (string?)mark.Attribute("textWrap"));
        var pos = mark.Element(Hp + "pos")!;
        Assert.Equal(("PAGE", "CENTER"), ((string?)pos.Attribute("horzRelTo"), (string?)pos.Attribute("horzAlign")));
        Assert.Equal(("PAGE", "CENTER"), ((string?)pos.Attribute("vertRelTo"), (string?)pos.Attribute("vertAlign")));

        using var package = ZipFile.OpenRead(hwpx);
        using var stream = package.Entries.Single(entry => entry.FullName.StartsWith("BinData/")).Open();
        using var image = new MagickImage(stream);
        Assert.InRange((double)image.Width / image.Height, 0.98, 1.02);
        Assert.True(image.HasAlpha);
        using var pixels = image.GetPixels();
        var rgba = pixels.ToByteArray(PixelMapping.RGBA)!;
        var maxAlpha = 0;
        for (var i = 3; i < rgba.Length; i += 4)
            maxAlpha = Math.Max(maxAlpha, rgba[i]);
        Assert.InRange(maxAlpha, 113, 116); // 29491 / 65536 * 255

        var renderer = fx.Converters.Single(c => c.Id == "rhwp");
        Assert.SkipUnless(fx.Catalog.StatusOf(renderer).IsAvailable, "rhwp not found");
        var source = docx.Save(Path.Combine(fx.NewFolder(), "source.docx"));
        var catalog = new ConverterCatalog(fx.Converters.Where(c => c.Id is "hwpx-writer" or "rhwp"));
        var pdf = await fx.ConvertAsync([source], new Preset { TargetFormat = "pdf" }, catalog);
        Assert.Equal(JobState.Completed, pdf.State);
        using var output = PdfDocument.Open(Assert.Single(pdf.Outputs));
        var page = Assert.Single(output.GetPages());
        var bounds = Assert.Single(page.GetImages()).BoundingBox;
        Assert.InRange((bounds.Top + bounds.Bottom) / 2, page.Height / 2 - 3, page.Height / 2 + 3);
    }
}
