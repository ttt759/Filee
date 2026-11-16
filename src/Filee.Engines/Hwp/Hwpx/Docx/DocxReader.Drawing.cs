// DOCX graphic objects: DrawingML pictures, text boxes (wps), groups and drawing canvases, plus legacy VML
// (w:pict / w:object) used by older documents and as the fallback of mc:AlternateContent.

using System.Globalization;
using System.Xml.Linq;
using Filee.Engines.Vector;
using ImageMagick;
using SkiaSharp;

namespace Filee.Engines.Hwp.Hwpx.Docx;

internal sealed partial class DocxReader
{
    private static readonly XNamespace Wpc = "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas";
    private static readonly XNamespace W10 = "urn:schemas-microsoft-com:office:word";

    /// <summary>Default DrawingML text box insets: 0.1" left/right, 0.05" top/bottom.</summary>
    private static readonly HInsets DefaultTextBoxPadding = new(720, 720, 360, 360);

    private readonly Dictionary<string, string> _themeColors = [];
    private readonly Dictionary<HParagraph, XElement> _frameProperties = [];

    /// <summary>A fixed-size paragraph frame carries its text and inline graphics together at a floating position.</summary>
    private bool TryReadFrame(XElement paragraph, List<HBlock> output)
    {
        var pPr = paragraph.Element(W + "pPr");
        var frame = pPr?.Element(W + "framePr");
        if (frame is null || pPr!.Element(W + "sectPr") is not null || Attr(frame, "dropCap") is "drop" or "margin"
            || Twips(frame, "w") is not { } width || width <= 0 || Twips(frame, "h") is not { } height || height <= 0
            || Twips(frame, "x") is not { } x || Twips(frame, "y") is not { } y)
            return false;

        if (output.LastOrDefault() is HParagraph previous && _frameProperties.TryGetValue(previous, out var properties)
            && XNode.DeepEquals(properties, frame) && previous.Inlines is [HTextBox existing])
        {
            ReadParagraph(paragraph, existing.Blocks);
            return true;
        }

        var anchor = new HAnchor(
            RelativeTo(Attr(frame, "hAnchor"), "COLUMN"), "LEFT", HwpxUnits.FromTwips(x),
            RelativeTo(Attr(frame, "vAnchor"), "PARA"), "TOP", HwpxUnits.FromTwips(y),
            Attr(frame, "wrap") switch
            {
                "around" or "auto" or "tight" or "through" => "SQUARE",
                "notBeside" => "TOP_AND_BOTTOM",
                _ => "IN_FRONT_OF_TEXT",
            }, true);
        var box = new TextBoxBuilder(HwpxUnits.FromTwips(width), HwpxUnits.FromTwips(height), anchor)
        {
            Line = HBorder.None,
            Padding = new HInsets(0, 0, 0, 0),
        }.Build();
        ReadParagraph(paragraph, box.Blocks);

        // The empty anchor stays in the body. The content and its line spacing belong to the frame.
        var first = (HParagraph)box.Blocks[0];
        var carrier = new HParagraph
        {
            PageBreakBefore = first.PageBreakBefore,
            ColumnBreakBefore = first.ColumnBreakBefore,
            Format = new HParaFormat(Before: 0, After: 0, LineSpacing: new HLineSpacing(HLineSpacingKind.Fixed, 1)),
        };
        first.PageBreakBefore = false;
        first.ColumnBreakBefore = false;
        carrier.Inlines.Add(box);
        _frameProperties[carrier] = frame;
        output.Add(carrier);
        return true;

        static string RelativeTo(string? value, string fallback) => value switch
        {
            "page" => "PAPER",
            "margin" => "PAGE",
            _ => fallback,
        };
    }

    private IEnumerable<HInline> Drawing(XElement drawing)
    {
        foreach (var container in drawing.Elements())
        {
            var anchored = container.Name == Wp + "anchor";
            if (!anchored && container.Name != Wp + "inline")
                continue;
            var extent = container.Element(Wp + "extent");
            var width = HwpxUnits.FromEmu(Long(extent, "cx") ?? 0);
            var height = HwpxUnits.FromEmu(Long(extent, "cy") ?? 0);
            var anchor = anchored ? Anchor(container) : null;
            var data = container.Element(A + "graphic")?.Element(A + "graphicData");
            foreach (var inline in GraphicObjects(data?.Elements() ?? [], width, height, anchor, (0, 0, 1, 1)))
                yield return inline;
        }
    }

    /// <summary>
    /// Pictures and text boxes of a graphic frame. Members of groups and canvases become separate objects placed at
    /// their position inside the group (<paramref name="groupTransform"/>: offset and scale of the child coordinates).
    /// </summary>
    private IEnumerable<HInline> GraphicObjects(IEnumerable<XElement> objects, int width, int height, HAnchor? anchor,
        (double X, double Y, double ScaleX, double ScaleY) groupTransform)
    {
        foreach (var obj in objects)
        {
            if (obj.Name == Pic + "pic")
            {
                if (Picture(obj, width, height, anchor) is { } picture)
                    yield return picture;
            }
            else if (obj.Name == Wps + "wsp")
            {
                if (Shape(obj, width, height, anchor) is { } shape)
                    yield return shape;
            }
            else if (obj.Name == Wpg + "wgp" || obj.Name == Wpg + "grpSp" || obj.Name == Wpc + "wpc")
            {
                foreach (var inline in Group(obj, width, height, anchor, groupTransform))
                    yield return inline;
            }
        }
    }

    private IEnumerable<HInline> Group(XElement group, int width, int height, HAnchor? anchor,
        (double X, double Y, double ScaleX, double ScaleY) parent)
    {
        var xfrm = group.Element(Wpg + "grpSpPr")?.Element(A + "xfrm") ?? group.Element(A + "grpSpPr")?.Element(A + "xfrm");
        var chOff = xfrm?.Element(A + "chOff");
        var chExt = xfrm?.Element(A + "chExt");
        // Child coordinates → HWPUNIT relative to the group's top-left corner.
        var scaleX = (Long(chExt, "cx") ?? 0) > 0 ? width / (double)Long(chExt, "cx")! : 1 / 127.0;
        var scaleY = (Long(chExt, "cy") ?? 0) > 0 ? height / (double)Long(chExt, "cy")! : 1 / 127.0;
        var originX = Long(chOff, "x") ?? 0;
        var originY = Long(chOff, "y") ?? 0;

        foreach (var child in group.Elements())
        {
            var childXfrm = child.Descendants(A + "xfrm").FirstOrDefault();
            var off = childXfrm?.Element(A + "off");
            var childExt = childXfrm?.Element(A + "ext");
            var x = (int)(((Long(off, "x") ?? 0) - originX) * scaleX);
            var y = (int)(((Long(off, "y") ?? 0) - originY) * scaleY);
            var childWidth = (int)((Long(childExt, "cx") ?? 0) * scaleX);
            var childHeight = (int)((Long(childExt, "cy") ?? 0) * scaleY);

            // An inline group becomes objects floating over the paragraph at the same place.
            var childAnchor = anchor is null
                ? new HAnchor("COLUMN", "LEFT", x, "PARA", "TOP", y, "IN_FRONT_OF_TEXT", true)
                : anchor with
                {
                    HorizontalOffset = anchor.HorizontalOffset + AlignedOffset(anchor.HorizontalAlign, x, childWidth, width),
                    VerticalOffset = anchor.VerticalOffset + AlignedOffset(anchor.VerticalAlign, y, childHeight, height),
                    AllowOverlap = true,
                };
            foreach (var inline in GraphicObjects([child], childWidth, childHeight, childAnchor, parent))
                yield return inline;
        }
    }

    /// <summary>
    /// Offset of a group member when the group is aligned (centered, right-aligned ...) rather than placed at an
    /// offset: the member keeps its distance from the group's center or far edge.
    /// </summary>
    private static int AlignedOffset(string align, int position, int size, int groupSize) => align switch
    {
        "CENTER" => position + size / 2 - groupSize / 2,
        "RIGHT" or "BOTTOM" => position + size - groupSize,
        _ => position,
    };

    private HImage? Picture(XElement pic, int width, int height, HAnchor? anchor)
    {
        var blip = pic.Descendants(A + "blip").FirstOrDefault();
        var path = ExtractImage(Attr(blip, "embed", R));
        if (path is not null && pic.Descendants(A + "srcRect").FirstOrDefault() is { } crop)
            path = CroppedImage(path, crop);
        return path is null ? null : new HImage(path) { Width = width > 0 ? width : null, Height = height > 0 ? height : null, Anchor = anchor };
    }

    /// <summary>A text box, or a rectangle / ellipse / line without text. Other geometries are skipped.</summary>
    private HInline? Shape(XElement wsp, int width, int height, HAnchor? anchor)
    {
        var spPr = wsp.Element(Wps + "spPr");
        var style = wsp.Element(Wps + "style");
        var (fill, gradient) = FillOf(spPr, style);
        var geometry = Attr(spPr?.Element(A + "prstGeom"), "prst");
        if (wsp.Element(Wps + "txbx")?.Element(W + "txbxContent") is { } content)
        {
            if (geometry is "rect" or "roundRect" or "flowChartProcess" or "flowChartAlternateProcess" or "ellipse" or "flowChartConnector")
            {
                var framed = TextBox(wsp, content, width, height, anchor, fill ?? gradient?.Colors[0]);
                if (geometry is "ellipse" or "flowChartConnector")
                    framed.Shape = HShapeKind.Ellipse;
                else if (geometry is "roundRect" or "flowChartAlternateProcess")
                    framed.CornerRatio = 20;
                return framed;
            }
            // Text on a free-form or decorative shape: keep the text, drop the outline 한글 cannot draw.
            if (string.IsNullOrWhiteSpace(content.Value))
                return null;
            var box = TextBox(wsp, content, width, height, anchor, fill: null);
            box.Line = HBorder.None;
            return box;
        }

        HShapeKind? kind = geometry switch
        {
            "rect" or "roundRect" or "snip1Rect" or "snip2SameRect" or "round1Rect" or "round2SameRect"
                or "flowChartProcess" or "flowChartAlternateProcess" or "frame" => HShapeKind.Rectangle,
            "ellipse" or "flowChartConnector" => HShapeKind.Ellipse,
            "line" or "straightConnector1" => HShapeKind.Line,
            _ => null, // free-form and other preset shapes have no 한글 equivalent
        };
        if (kind is null)
            return null;
        var xfrm = spPr?.Element(A + "xfrm");
        return new HShape(kind.Value)
        {
            Width = width,
            Height = height,
            Anchor = anchor,
            Fill = kind == HShapeKind.Line ? null : fill,
            Gradient = kind == HShapeKind.Line ? null : gradient,
            Line = Line(spPr?.Element(A + "ln"), style),
            CornerRatio = geometry is "roundRect" or "round1Rect" or "round2SameRect" or "flowChartAlternateProcess" ? 20 : 0,
            FlipHorizontal = Attr(xfrm, "flipH") is "1" or "true",
            FlipVertical = Attr(xfrm, "flipV") is "1" or "true",
        };
    }

    private HTextBox TextBox(XElement wsp, XElement content, int width, int height, HAnchor? anchor, string? fill)
    {
        var spPr = wsp.Element(Wps + "spPr");
        var bodyPr = wsp.Element(Wps + "bodyPr");
        var box = new TextBoxBuilder(width, height, anchor)
        {
            Fill = fill,
            Line = Line(spPr?.Element(A + "ln"), wsp.Element(Wps + "style")),
            Padding = new HInsets(
                Inset(bodyPr, "lIns", DefaultTextBoxPadding.Left), Inset(bodyPr, "rIns", DefaultTextBoxPadding.Right),
                Inset(bodyPr, "tIns", DefaultTextBoxPadding.Top), Inset(bodyPr, "bIns", DefaultTextBoxPadding.Bottom)),
            VerticalAlign = Attr(bodyPr, "anchor") switch
            {
                "ctr" => HVerticalAlign.Center,
                "b" => HVerticalAlign.Bottom,
                _ => HVerticalAlign.Top,
            },
        }.Build();
        ReadTextBoxContent(content, box.Blocks);
        return box;
    }

    /// <summary>
    /// Text box paragraphs; Word's document grid does not apply inside text boxes, nor does the style of a table the
    /// text box sits in.
    /// </summary>
    private void ReadTextBoxContent(XElement content, List<HBlock> output)
    {
        _textBoxDepth++;
        var tableStyleLayer = _tableStyleLayer;
        _tableStyleLayer = null;
        try
        {
            ReadNestedBlocks(content, output);
        }
        finally
        {
            _textBoxDepth--;
            _tableStyleLayer = tableStyleLayer;
        }
    }

    /// <summary>Collects text box settings before the HTextBox (whose Blocks list is read-only) is created.</summary>
    private sealed class TextBoxBuilder(int width, int height, HAnchor? anchor)
    {
        public string? Fill { get; init; }
        public HBorder Line { get; init; } = HBorder.Thin;
        public HInsets Padding { get; init; } = DefaultTextBoxPadding;
        public HVerticalAlign VerticalAlign { get; init; }

        public HTextBox Build() => new()
        {
            Width = Math.Max(width, 1000),
            Height = Math.Max(height, 600),
            Anchor = anchor,
            Fill = Fill,
            Line = Line,
            Padding = Padding,
            VerticalAlign = VerticalAlign,
        };
    }

    private static int Inset(XElement? bodyPr, string name, int fallback) =>
        Long(bodyPr, name) is { } emu ? HwpxUnits.FromEmu(emu) : fallback;

    /// <summary>Solid colour or linear gradient of a shape; without its own fill the shape style decides.</summary>
    private (string? Solid, HGradient? Gradient) FillOf(XElement? spPr, XElement? style)
    {
        if (spPr?.Element(A + "noFill") is not null)
            return (null, null);
        if (spPr?.Element(A + "solidFill") is { } solid)
            return (Color(solid), null);
        if (spPr?.Element(A + "gradFill") is { } gradient)
        {
            var colors = gradient.Element(A + "gsLst")?.Elements(A + "gs")
                .OrderBy(gs => Long(gs, "pos") ?? 0)
                .Select(Color)
                .OfType<string>()
                .ToList() ?? [];
            var angle = (int)((Long(gradient.Element(A + "lin"), "ang") ?? 0) / 60000);
            return colors.Count >= 2 ? (null, new HGradient(colors, angle)) : (colors.FirstOrDefault(), null);
        }
        var fillRef = style?.Element(A + "fillRef");
        return fillRef is not null && Attr(fillRef, "idx") is not (null or "0") ? (Color(fillRef), null) : (null, null);
    }

    private HBorder Line(XElement? ln, XElement? style)
    {
        if (ln is null)
        {
            // No explicit outline: the shape style decides (lnRef idx 0 = no line).
            var lnRef = style?.Element(A + "lnRef");
            return lnRef is not null && Attr(lnRef, "idx") is not (null or "0")
                ? HBorder.Thin with { Color = Color(lnRef) ?? "#000000" }
                : HBorder.None;
        }
        if (ln.Element(A + "noFill") is not null)
            return HBorder.None;
        var widthMm = Long(ln, "w") is { } emu ? emu / 36000.0 : 0.12;
        var dash = Attr(ln.Element(A + "prstDash"), "val");
        var borderStyle = dash switch
        {
            null or "solid" => HBorderStyle.Solid,
            "dot" or "sysDot" => HBorderStyle.Dot,
            "dashDot" or "sysDashDot" or "lgDashDot" => HBorderStyle.DashDot,
            _ => HBorderStyle.Dash,
        };
        return new HBorder(borderStyle, widthMm, Color(ln.Element(A + "solidFill")) ?? "#000000");
    }

    /// <summary>Colour of a fill element: srgbClr, sysClr or a theme colour (luminance modifiers are ignored).</summary>
    private string? Color(XElement? fill)
    {
        if (fill is null)
            return null;
        if (fill.Element(A + "srgbClr") is { } rgb)
            return "#" + (Val(rgb) ?? "000000").ToUpperInvariant();
        if (fill.Element(A + "sysClr") is { } sys)
            return "#" + (Attr(sys, "lastClr") ?? "000000").ToUpperInvariant();
        if (fill.Element(A + "schemeClr") is { } scheme)
        {
            var name = Val(scheme) switch
            {
                "bg1" => "lt1",
                "tx1" => "dk1",
                "bg2" => "lt2",
                "tx2" => "dk2",
                var other => other ?? "dk1",
            };
            return _themeColors.GetValueOrDefault(name, name.StartsWith("lt", StringComparison.Ordinal) ? "#FFFFFF" : "#000000");
        }
        return null;
    }

    private void LoadThemeColors(Part? theme)
    {
        foreach (var color in theme?.Xml.Descendants(A + "clrScheme").FirstOrDefault()?.Elements() ?? [])
        {
            var value = Val(color.Element(A + "srgbClr")) ?? Attr(color.Element(A + "sysClr"), "lastClr");
            if (value is { Length: 6 })
                _themeColors[color.Name.LocalName] = "#" + value.ToUpperInvariant();
        }
    }

    /// <summary>Position of a floating DrawingML object (wp:anchor).</summary>
    private static HAnchor Anchor(XElement anchor)
    {
        var horizontal = anchor.Element(Wp + "positionH");
        var vertical = anchor.Element(Wp + "positionV");
        var wrap = anchor.Elements().FirstOrDefault(e => e.Name.Namespace == Wp && e.Name.LocalName.StartsWith("wrap", StringComparison.Ordinal))?.Name.LocalName;
        var behind = Attr(anchor, "behindDoc") is "1" or "true";
        return new HAnchor(
            HorizontalRelativeTo: Attr(horizontal, "relativeFrom") switch
            {
                "page" or "leftMargin" or "rightMargin" or "insideMargin" or "outsideMargin" => "PAPER",
                "margin" => "PAGE",
                "character" => "PARA",
                _ => "COLUMN",
            },
            HorizontalAlign: horizontal?.Element(Wp + "align")?.Value switch
            {
                "center" => "CENTER",
                "right" => "RIGHT",
                "inside" => "INSIDE",
                "outside" => "OUTSIDE",
                _ => "LEFT",
            },
            HorizontalOffset: OffsetOf(horizontal),
            VerticalRelativeTo: Attr(vertical, "relativeFrom") switch
            {
                "page" or "topMargin" or "bottomMargin" or "insideMargin" or "outsideMargin" => "PAPER",
                "margin" => "PAGE",
                _ => "PARA",
            },
            VerticalAlign: vertical?.Element(Wp + "align")?.Value switch
            {
                "center" => "CENTER",
                "bottom" => "BOTTOM",
                "inside" => "INSIDE",
                "outside" => "OUTSIDE",
                _ => "TOP",
            },
            VerticalOffset: OffsetOf(vertical),
            Wrap: wrap switch
            {
                "wrapSquare" or "wrapTight" or "wrapThrough" => "SQUARE",
                "wrapTopAndBottom" => "TOP_AND_BOTTOM",
                _ => behind ? "BEHIND_TEXT" : "IN_FRONT_OF_TEXT",
            },
            AllowOverlap: Attr(anchor, "allowOverlap") is not ("0" or "false"));
    }

    private static int OffsetOf(XElement? position) =>
        long.TryParse(position?.Element(Wp + "posOffset")?.Value, out var emu) ? HwpxUnits.FromEmu(emu) : 0;

    private string? ExtractImage(string? relationshipId)
    {
        if (relationshipId is null || !_part.Relationships.TryGetValue(relationshipId, out var rel) || rel.External)
            return null;
        if (_media.TryGetValue(rel.Target, out var existing))
            return existing;
        if (_zip.GetEntry(rel.Target) is not { } entry)
            return null;
        Directory.CreateDirectory(_mediaFolder);
        var path = Path.Combine(_mediaFolder, $"docx-media-{_media.Count + 1}{Path.GetExtension(rel.Target)}");
        _unpack.Extract(entry, path);
        _media[rel.Target] = path;
        return path;
    }

    /// <summary>
    /// A copy of a picture with a:srcRect's edges cut off (each in 1/1000 percent of the picture), or the picture itself
    /// when nothing is cut or it can't be read (EMF). Word keeps the whole picture in the file and shows the rest; a
    /// negative edge adds space around it, which is not done here.
    /// </summary>
    private string CroppedImage(string path, XElement crop)
    {
        static double Edge(XElement rect, string name) =>
            int.TryParse((string?)rect.Attribute(name), out var value) ? Math.Clamp(value, 0, 100_000) / 100_000.0 : 0;
        var (left, top, right, bottom) = (Edge(crop, "l"), Edge(crop, "t"), Edge(crop, "r"), Edge(crop, "b"));
        if (left + top + right + bottom == 0 || left + right >= 1 || top + bottom >= 1)
            return path;
        var key = $"{path}|{left}|{top}|{right}|{bottom}";
        if (_media.TryGetValue(key, out var existing))
            return existing;
        try
        {
            using var image = new MagickImage(path);
            var (width, height) = ((double)image.Width, (double)image.Height);
            image.Crop(new MagickGeometry(
                (int)Math.Round(width * left), (int)Math.Round(height * top),
                (uint)Math.Max(1, Math.Round(width * (1 - left - right))), (uint)Math.Max(1, Math.Round(height * (1 - top - bottom)))));
            image.ResetPage();
            var cropped = Path.Combine(_mediaFolder, $"docx-media-{_media.Count + 1}{Path.GetExtension(path)}");
            image.Write(cropped);
            _media[key] = cropped;
            return cropped;
        }
        catch (MagickException)
        {
            return path;
        }
    }

    // ───────────────────────── VML (legacy) ─────────────────────────

    /// <param name="format">Formatting of the run holding the objects (WordArt text starts from it).</param>
    private IEnumerable<HInline> Vml(XElement container, HCharFormat format)
    {
        foreach (var shape in container.Elements())
        {
            if (shape.Name.Namespace != V)
                continue;
            if (shape.Name.LocalName == "group")
            {
                foreach (var inline in Vml(shape, format))
                    yield return inline;
                continue;
            }
            if (shape.Name.LocalName is "shapetype")
                continue;

            var style = VmlStyle((string?)shape.Attribute("style"));
            if (style.GetValueOrDefault("visibility") == "hidden")
                continue;
            var width = HwpxUnits.ParseLength(style.GetValueOrDefault("width")) ?? 0;
            var height = HwpxUnits.ParseLength(style.GetValueOrDefault("height")) ?? 0;
            var anchor = style.GetValueOrDefault("position") == "absolute" ? VmlAnchor(shape, style) : null;

            if (shape.Element(V + "imagedata") is { } imageData && ExtractImage(Attr(imageData, "id", R)) is { } path)
            {
                yield return new HImage(path) { Width = width > 0 ? width : null, Height = height > 0 ? height : null, Anchor = anchor };
            }
            else if (shape.Element(V + "textbox")?.Element(W + "txbxContent") is { } content)
            {
                var box = new TextBoxBuilder(width, height, anchor)
                {
                    Fill = VmlFill(shape),
                    Line = VmlLine(shape),
                }.Build();
                ReadTextBoxContent(content, box.Blocks);
                yield return box;
            }
            else if (WordArt(shape, width, height, format) is { } text)
            {
                if (anchor is null)
                {
                    yield return text;
                }
                else if (anchor.Wrap == "BEHIND_TEXT" && WordArtWatermark(shape, style, text, width, height, anchor) is { } watermark)
                {
                    yield return watermark;
                }
                else
                {
                    var box = new TextBoxBuilder(width, height, anchor) { Line = HBorder.None, Padding = new HInsets(0, 0, 0, 0), VerticalAlign = HVerticalAlign.Center }.Build();
                    var paragraph = new HParagraph { Format = new HParaFormat(Align: HAlign.Center) };
                    paragraph.Inlines.Add(text);
                    box.Blocks.Add(paragraph);
                    yield return box;
                }
            }
            else if (shape.Name.LocalName is "rect" or "roundrect" or "oval" or "line")
            {
                var kind = shape.Name.LocalName switch
                {
                    "oval" => HShapeKind.Ellipse,
                    "line" => HShapeKind.Line,
                    _ => HShapeKind.Rectangle,
                };
                if (kind == HShapeKind.Line)
                {
                    // v:line has from="x,y" to="x,y" instead of a size.
                    var from = VmlPoint((string?)shape.Attribute("from"));
                    var to = VmlPoint((string?)shape.Attribute("to"));
                    width = Math.Abs(to.X - from.X);
                    height = Math.Abs(to.Y - from.Y);
                    var lineAnchor = anchor ?? VmlAnchor(shape, style);
                    anchor = lineAnchor with
                    {
                        HorizontalOffset = lineAnchor.HorizontalOffset + Math.Min(from.X, to.X),
                        VerticalOffset = lineAnchor.VerticalOffset + Math.Min(from.Y, to.Y),
                    };
                }
                yield return new HShape(kind)
                {
                    Width = width,
                    Height = height,
                    Anchor = anchor,
                    Fill = kind == HShapeKind.Line ? null : VmlFill(shape),
                    Line = VmlLine(shape),
                    CornerRatio = shape.Name.LocalName == "roundrect" ? 20 : 0,
                };
            }
        }
    }

    /// <summary>
    /// Legacy WordArt (a VML shape with v:textpath, as WPS and older Word write it) as plain text in its font and
    /// fill colour. 한글 has no text effects, so the text keeps the run's other formatting; when the text is stretched
    /// to the shape (fitpath / fitshape), the font size is chosen so it fills about the same box.
    /// </summary>
    private static HText? WordArt(XElement shape, int width, int height, HCharFormat format)
    {
        var path = shape.Element(V + "textpath");
        var text = (string?)path?.Attribute("string");
        if (path is null || (string?)path.Attribute("on") is "f" or "false" || string.IsNullOrWhiteSpace(text))
            return null;

        // An empty move followed by an empty line has no baseline for Word to draw the text on.
        if (Attr(shape, "path")?.StartsWith("m,l,", StringComparison.Ordinal) == true)
            return null;
        var font = VmlStyle((string?)path.Attribute("style"));
        var family = font.GetValueOrDefault("font-family")?.Split(',', 2)[0].Trim('"', '\'', ' ');
        var size = HwpxUnits.ParseLength(font.GetValueOrDefault("font-size")) is { } points and > 0 ? points : format.Size ?? 1000;
        var stretched = (string?)path.Attribute("fitpath") is "t" or "true" || (string?)path.Attribute("fitshape") is "t" or "true";
        if (stretched && width > 0 && height > 0)
        {
            // The glyphs fill the shape: capitals and digits are about 0.72 em tall, Han characters about 0.9 em;
            // East Asian characters are 1 em wide, others about half.
            var wide = text.Count(c => c is >= '⺀' and <= '鿿' or >= '가' and <= '힣' or >= '＀' and <= '｠');
            var ems = wide + (text.Length - wide) * 0.5;
            var byHeight = height / (wide > 0 ? 0.9 : 0.72);
            size = (int)Math.Min(byHeight, width / Math.Max(ems, 1));
        }
        var color = (string?)shape.Attribute("filled") is "f" or "false"
            ? VmlColor((string?)shape.Attribute("strokecolor"))
            : VmlColor((string?)shape.Attribute("fillcolor"));
        return new HText(text.Replace('\n', ' ').Replace('\r', ' '), format with
        {
            Size = Math.Max(size, 100),
            Font = family ?? format.Font,
            EastAsianFont = family ?? format.EastAsianFont,
            Bold = font.GetValueOrDefault("font-weight") is { } weight ? weight == "bold" || int.TryParse(weight, out var w) && w >= 600 : format.Bold,
            Italic = font.GetValueOrDefault("font-style") is { } italic ? italic == "italic" : format.Italic,
            Color = color ?? format.Color,
        });
    }

    /// <summary>
    /// Draws a watermark as a translucent graphic. HWPX readers may rotate the outline of a text box and leave its
    /// text horizontal. The rendered glyphs fit the original VML box and turn around its center.
    /// </summary>
    private HImage? WordArtWatermark(XElement shape, Dictionary<string, string> style, HText text, int width, int height, HAnchor anchor)
    {
        var angle = VmlNumber(style.GetValueOrDefault("rotation"), "fd", 0) % 360;
        var opacityAttribute = Attr(shape.Element(V + "fill"), "opacity") ?? Attr(shape, "opacity");
        var opacity = Math.Clamp(VmlNumber(opacityAttribute, "f", 1), 0, 1);
        if (width <= 0 || height <= 0 || (angle == 0 && opacityAttribute is null))
            return null;
        var color = SKColor.TryParse(text.Format.Color, out var fill) ? fill : SKColors.Black;
        var alpha = (byte)Math.Round(opacity * 255);

        var fontStyle = text.Format.Bold == true
            ? (text.Format.Italic == true ? SKFontStyle.BoldItalic : SKFontStyle.Bold)
            : (text.Format.Italic == true ? SKFontStyle.Italic : SKFontStyle.Normal);
        var other = text.Text.EnumerateRunes().FirstOrDefault(r => r.Value > 0x24f);
        // SkiaSharp can return the same managed typeface from multiple font lookups.
        var requested = SKFontManager.Default.MatchFamily(text.Format.Font, fontStyle);
        var fallback = requested?.ContainsGlyphs(text.Text) == true || other.Value == 0
            ? null : SKFontManager.Default.MatchCharacter(null, fontStyle, null, other.Value);
        using var font = new SKFont(fallback ?? requested ?? SKTypeface.Default, 100);
        using var glyphs = new GlyphOutlines();
        using var outline = new SKPath { FillType = SKPathFillType.Winding };
        if (!glyphs.AppendShaped(outline, text.Text, font, 0, 0, SKTextAlign.Left) || outline.IsEmpty)
            return null;
        var bounds = outline.TightBounds;
        if (bounds.Width <= 0 || bounds.Height <= 0)
            return null;

        var radians = angle * Math.PI / 180;
        var (cosine, sine) = (Math.Abs(Math.Cos(radians)), Math.Abs(Math.Sin(radians)));
        var rotatedWidth = (int)Math.Min(int.MaxValue, Math.Ceiling(width * cosine + height * sine));
        var rotatedHeight = (int)Math.Min(int.MaxValue, Math.Ceiling(width * sine + height * cosine));
        var scale = (float)Math.Min(0.03, 2400.0 / Math.Max(rotatedWidth, rotatedHeight)); // 216 dpi, bounded
        var pixelsWide = (int)Math.Ceiling(rotatedWidth * scale);
        var pixelsHigh = (int)Math.Ceiling(rotatedHeight * scale);
        using var bitmap = new SKBitmap(pixelsWide, pixelsHigh, SKColorType.Rgba8888, SKAlphaType.Premul);
        using var canvas = new SKCanvas(bitmap);
        using var paint = new SKPaint { Color = color.WithAlpha(alpha), IsAntialias = true };
        canvas.Clear(SKColors.Transparent);
        canvas.Translate(pixelsWide / 2f, pixelsHigh / 2f);
        canvas.RotateDegrees((float)angle);
        canvas.Scale(width * scale / bounds.Width, height * scale / bounds.Height);
        canvas.Translate(-bounds.MidX, -bounds.MidY);
        canvas.DrawPath(outline, paint);

        var key = $"watermark|{width}|{height}|{angle}|{opacity}|{text.Format}|{text.Text}";
        if (!_media.TryGetValue(key, out var path))
        {
            Directory.CreateDirectory(_mediaFolder);
            path = Path.Combine(_mediaFolder, $"docx-media-{_media.Count + 1}.png");
            using var data = bitmap.Encode(SKEncodedImageFormat.Png, 100);
            using var file = File.Create(path);
            data.SaveTo(file);
            _media[key] = path;
        }
        return new HImage(path)
        {
            Width = rotatedWidth,
            Height = rotatedHeight,
            Anchor = anchor with
            {
                HorizontalOffset = anchor.HorizontalOffset + AlignedOffset(anchor.HorizontalAlign, (width - rotatedWidth) / 2, rotatedWidth, width),
                VerticalOffset = anchor.VerticalOffset + AlignedOffset(anchor.VerticalAlign, (height - rotatedHeight) / 2, rotatedHeight, height),
            },
        };
    }

    private static double VmlNumber(string? value, string fixedSuffix, double fallback)
    {
        if (value is null)
            return fallback;
        var divisor = value.EndsWith(fixedSuffix, StringComparison.Ordinal) ? 65536.0 : value.EndsWith('%') ? 100 : 1;
        var number = divisor == 65536 ? value[..^fixedSuffix.Length] : divisor == 100 ? value[..^1] : value;
        return double.TryParse(number, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed) && double.IsFinite(parsed)
            ? parsed / divisor : fallback;
    }

    private static string? VmlFill(XElement shape) =>
        (string?)shape.Attribute("filled") is "f" or "false" ? null : VmlColor((string?)shape.Attribute("fillcolor")) ?? "#FFFFFF";

    private static HBorder VmlLine(XElement shape) =>
        (string?)shape.Attribute("stroked") is "f" or "false"
            ? HBorder.None
            : HBorder.Thin with
            {
                Color = VmlColor((string?)shape.Attribute("strokecolor")) ?? "#000000",
                WidthMm = HwpxUnits.ParseLength((string?)shape.Attribute("strokeweight")) is { } weight ? weight / HwpxUnits.PerMm : 0.12,
            };

    /// <summary>"x,y" of a VML line end; bare numbers are points.</summary>
    private static (int X, int Y) VmlPoint(string? point)
    {
        var parts = (point ?? "0,0").Split(',');
        int Coordinate(int index)
        {
            var text = parts.ElementAtOrDefault(index)?.Trim() ?? "0";
            if (text.Length == 0)
                text = "0";
            return HwpxUnits.ParseLength(char.IsDigit(text[^1]) ? text + "pt" : text) ?? 0;
        }
        return (Coordinate(0), Coordinate(1));
    }

    private static Dictionary<string, string> VmlStyle(string? style) =>
        (style ?? "").Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(p => p.Split(':', 2))
            .Where(p => p.Length == 2)
            .GroupBy(p => p[0].Trim().ToLowerInvariant())
            .ToDictionary(g => g.Key, g => g.Last()[1].Trim());

    private static HAnchor VmlAnchor(XElement shape, Dictionary<string, string> style)
    {
        var wrap = (string?)shape.Element(W10 + "wrap")?.Attribute("type");
        return new HAnchor(
            HorizontalRelativeTo: style.GetValueOrDefault("mso-position-horizontal-relative") switch
            {
                "page" => "PAPER",
                "margin" => "PAGE",
                "char" => "PARA",
                _ => "COLUMN",
            },
            HorizontalAlign: style.GetValueOrDefault("mso-position-horizontal") switch
            {
                "center" => "CENTER",
                "right" => "RIGHT",
                _ => "LEFT",
            },
            HorizontalOffset: HwpxUnits.ParseLength(style.GetValueOrDefault("margin-left")) ?? 0,
            VerticalRelativeTo: style.GetValueOrDefault("mso-position-vertical-relative") switch
            {
                "page" => "PAPER",
                "margin" => "PAGE",
                _ => "PARA",
            },
            VerticalAlign: style.GetValueOrDefault("mso-position-vertical") switch
            {
                "center" => "CENTER",
                "bottom" => "BOTTOM",
                _ => "TOP",
            },
            VerticalOffset: HwpxUnits.ParseLength(style.GetValueOrDefault("margin-top")) ?? 0,
            Wrap: wrap switch
            {
                "square" or "tight" or "through" => "SQUARE",
                "topAndBottom" => "TOP_AND_BOTTOM",
                _ => style.TryGetValue("z-index", out var z) && z.StartsWith('-') ? "BEHIND_TEXT" : "IN_FRONT_OF_TEXT",
            },
            AllowOverlap: true);
    }

    /// <summary>VML colours: "#RRGGBB", "#RGB" or "#RRGGBB [index]".</summary>
    private static string? VmlColor(string? value)
    {
        var color = value?.Split(' ')[0].Trim();
        if (color is null || !color.StartsWith('#'))
            return null;
        if (color.Length == 4)
            color = $"#{color[1]}{color[1]}{color[2]}{color[2]}{color[3]}{color[3]}";
        return color.Length == 7 && int.TryParse(color.AsSpan(1), NumberStyles.HexNumber, null, out _) ? color.ToUpperInvariant() : null;
    }
}
