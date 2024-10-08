using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using D = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DP = DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using System.IO;
using System.IO.Compression;
using XMLyzeLibrary.Interpreter;
using System.Globalization;
using System.Text.RegularExpressions;

namespace XMLyzeLibrary.Word
{
    public static class WF
    {
        public static (MainDocumentPart, Body) PopulateNewWordPackage(
            WordprocessingDocument package,
            List<Style> styleList,
            UInt32Value? margin = null,
            string? border = null)
        {
            if (margin == null) margin = 1440;

            // Create document structure in new package
            MainDocumentPart mainPart = package.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());

            // Add section properties & set page margins
            SectionProperties sectionProperties = new(
                new SectionType() { Val = SectionMarkValues.Continuous },
                new PageMargin()
                {
                    Top = new Int32Value((int)margin.Value),
                    Right = margin,
                    Bottom = new Int32Value((int)margin.Value),
                    Left = margin
                }
            );

            // Page border
            if (border == "blue")
                sectionProperties.Append(PageBorders(BorderValues.ThinThickThinSmallGap, 24, 16, "95DCF7", ThemeColorValues.Accent4, "66"));

            body.Append(sectionProperties);

            // Numbering definitions
            NumberingDefinitionsPart numberingPart = mainPart.AddNewPart<NumberingDefinitionsPart>("NumberingDefinitionsPart");
            numberingPart.Numbering = new();

            // Add styles
            StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
            Styles styles = new(styleList);
            styles.Save(stylePart);

            return (mainPart, body);
        }

        public static PageBorders PageBorders(EnumValue<BorderValues> value, UInt32Value size, UInt32Value space, string color, EnumValue<ThemeColorValues> themeColor, string themeTint)
        {
            PageBorders pageBorders = new(
                new TopBorder()
                {
                    Val = value,
                    Size = size,
                    Space = space,
                    Color = color,
                    ThemeColor = themeColor,
                    ThemeTint = themeTint
                },
                new LeftBorder()
                {
                    Val = value,
                    Size = size,
                    Space = space,
                    Color = color,
                    ThemeColor = themeColor,
                    ThemeTint = themeTint
                },
                new BottomBorder()
                {
                    Val = value,
                    Size = size,
                    Space = space,
                    Color = color,
                    ThemeColor = themeColor,
                    ThemeTint = themeTint
                },
                new RightBorder()
                {
                    Val = value,
                    Size = size,
                    Space = space,
                    Color = color,
                    ThemeColor = themeColor,
                    ThemeTint = themeTint
                }
            );

            pageBorders.OffsetFrom = PageBorderOffsetValues.Page;

            return pageBorders;
        }

        public static void AppendToBody(Body body, OpenXmlElement element)
        {
            SectionProperties? finalSectionProps = body.Elements<SectionProperties>().LastOrDefault();
            if (finalSectionProps != null)
                finalSectionProps.InsertBeforeSelf(element);
            else
                body.Append(element);
        }

        public static void AppendToBody(Body body, List<OpenXmlElement> elements)
        {
            SectionProperties? finalSectionProps = body.Elements<SectionProperties>().LastOrDefault();
            if (finalSectionProps != null)
                foreach (OpenXmlElement element in elements)
                    finalSectionProps.InsertBeforeSelf(element);
            else
                body.Append(elements);
        }

        public static void AppendToBody(Body body, List<Paragraph> paragraphs)
        {
            SectionProperties? finalSectionProps = body.Elements<SectionProperties>().LastOrDefault();
            if (finalSectionProps != null)
                foreach (Paragraph paragraph in paragraphs)
                    finalSectionProps.InsertBeforeSelf(paragraph);
            else
                body.Append(paragraphs);
        }

        public static Style Style(
            string id,
            string name,
            string? parentStyle = null,
            ParagraphProperties? pPr = null,
            StyleRunProperties? rPr = null,
            TableProperties? tblPr = null
        )
        {
            Style style = new(
                new AutoRedefine() { Val = OnOffOnlyValues.Off },
                new BasedOn() { Val = "Normal" },
                new LinkedStyle() { Val = "OverdueAmountChar" },
                new Locked() { Val = OnOffOnlyValues.Off },
                new PrimaryStyle() { Val = OnOffOnlyValues.On },
                new StyleHidden() { Val = OnOffOnlyValues.Off },
                new SemiHidden() { Val = OnOffOnlyValues.Off },
                new StyleName() { Val = name },
                new NextParagraphStyle() { Val = "Normal" },
                new UIPriority() { Val = 1 },
                new UnhideWhenUsed() { Val = OnOffOnlyValues.On }
            )
            {
                Type = tblPr != null ? StyleValues.Table : StyleValues.Paragraph,
                StyleId = id,
                CustomStyle = true,
                Default = false
            };

            if (parentStyle != null)
                style.AppendChild(new BasedOn() { Val = parentStyle });

            if (pPr != null)
                style.AppendChild(pPr);
            if (rPr != null)
                style.AppendChild(rPr);
            if (tblPr != null)
                style.AppendChild(tblPr);

            return style;
        }

        // Receives a list of arguments from a code block with a style command
        public static Style Style(List<IF.Argument> args)
        {
            string? name = null;
            string? parent = null;
            string color = "000000";
            string size = "24";
            string font = "Aptos";

            foreach (IF.Argument arg in args)
            {
                switch (arg.Name)
                {
                    case "name":
                        name = arg.Value;
                        break;
                    case "parent":
                        parent = arg.Value;
                        break;
                    case "color":
                        color = arg.Value;
                        break;
                    case "size":
                        size = $"{int.Parse(arg.Value) * 2}";
                        break;
                    case "font":
                        font = arg.Value;
                        break;
                    default:
                        throw new Exception($"Cannot recognize argument called {arg.Value}");
                }
            }

            // Check name and create id
            if (string.IsNullOrEmpty(name))
                throw new Exception("Style must have a name");
            string id = ToPascalCase(name);

            // Check color
            if (string.IsNullOrEmpty(color) && !IsValidHexCode(color))
                throw new Exception($"Color argument has an invalid hex code: {color}");
            if (color.Length == 3)
                color = ConvertToSixDigitHex(color);


            return Style(
                id,
                name,
                parent,
                new ParagraphProperties(
                    new SpacingBetweenLines()
                    {
                        Line = "276",
                        LineRule = LineSpacingRuleValues.Auto,
                        Before = "0",
                        After = "0"
                    }
                ),
                new StyleRunProperties(
                    new Color() { Val = color },
                    new RunFonts() { Ascii = font },
                    new FontSize() { Val = size },
                    new FontSizeComplexScript() { Val = size }
                )
            );
        }

        public static string ToPascalCase(string input)
        {
            // Split the input string into words by non-letter characters
            string[] words = Regex.Split(input, @"[^a-zA-Z0-9]+");

            // Capitalize the first letter of each word and join them
            TextInfo textInfo = CultureInfo.InvariantCulture.TextInfo;
            for (int i = 0; i < words.Length; i++)
                if (words[i].Length > 0)
                    words[i] = textInfo.ToTitleCase(words[i].ToLower());

            return string.Join(string.Empty, words);
        }

        private static bool IsValidHexCode(string? input)
        {
            if (input == null) return false;

            // Regular expression to match exactly 6 hexadecimal characters (0-9, A-F, a-f)
            Regex hexRegex = new Regex(@"^[0-9A-Fa-f]{6}$");

            // Returns true if the input matches the hex format
            return hexRegex.IsMatch(input);
        }

        private static string ConvertToSixDigitHex(string hex3)
        {
            if (hex3.Length != 3) throw new ArgumentException("Input must be a 3-digit hex code.");

            return string.Concat(
                hex3[0], hex3[0],
                hex3[1], hex3[1],
                hex3[2], hex3[2]
            );
        }

        public static Paragraph Paragraph(string text = "", string styleName = "Normal")
        {
            return new Paragraph(
                ParagraphStyle(ToPascalCase(styleName)),
                new Run(new Text(text))
            );
        }

        public static ParagraphProperties ParagraphStyle(string styleId)
        {
            return new ParagraphProperties(
                new ParagraphStyleId() { Val = styleId }
            );
        }

        public static Paragraph Image(
            string fileName,
            string relationshipId,
            Int64Value width,
            Int64Value height,
            string? styleId = null,
            bool rounded = false)
        {
            Paragraph paragraph = new(new Run(
                Drawing(fileName, relationshipId, width, height, rounded)
            ));

            if (styleId != null)
                paragraph.PrependChild(ParagraphStyle(styleId));

            return paragraph;
        }

        private static UInt32Value docPropertiesId = 0;

        public static Drawing Drawing(string fileName, string relationshipId, Int64Value width, Int64Value height, bool rounded = false)
        {
            docPropertiesId++;

            return new Drawing(
                new DW.Inline(
                    new DW.Extent() { Cx = width, Cy = height },
                    new DW.EffectExtent()
                    {
                        LeftEdge = 0L,
                        TopEdge = 0L,
                        RightEdge = 0L,
                        BottomEdge = 0L
                    },
                    new DW.DocProperties()
                    {
                        Id = docPropertiesId,
                        Name = fileName
                    },
                    new DW.NonVisualGraphicFrameDrawingProperties(
                        new D.GraphicFrameLocks() { NoChangeAspect = true }),
                    new D.Graphic(
                        new D.GraphicData(
                            new DP.Picture(
                                new DP.NonVisualPictureProperties(
                                    new DP.NonVisualDrawingProperties()
                                    {
                                        Id = docPropertiesId,
                                        Name = fileName
                                    },
                                    new DP.NonVisualPictureDrawingProperties()),
                                new DP.BlipFill(
                                    new D.Blip(
                                        new D.BlipExtensionList(
                                            new D.BlipExtension()
                                            {
                                                Uri =
                                                "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                            })
                                    )
                                    {
                                        Embed = relationshipId,
                                        CompressionState = D.BlipCompressionValues.Print
                                    },
                                    new D.Stretch(
                                        new D.FillRectangle())),
                                new DP.ShapeProperties(
                                    new D.Transform2D(
                                        new D.Offset() { X = 0L, Y = 0L },
                                        new D.Extents() { Cx = width, Cy = height }),
                                    new D.PresetGeometry(
                                        new D.AdjustValueList()
                                    )
                                    { Preset = rounded ? D.ShapeTypeValues.RoundRectangle : D.ShapeTypeValues.Rectangle }))
                        )
                        { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                )
                {
                    DistanceFromTop = (UInt32Value)0U,
                    DistanceFromBottom = (UInt32Value)0U,
                    DistanceFromLeft = (UInt32Value)0U,
                    DistanceFromRight = (UInt32Value)0U,
                    EditId = RandomEditId()
                });
        }

        public static string RandomEditId()
        {
            // Generate a random hexadecimal string
            byte[] bytes = new byte[4];
            new Random().NextBytes(bytes);

            // Convert byte array to hexadecimal string
            string randomHex = BitConverter.ToString(bytes).Replace("-", "");

            // Ensure the string has exactly 8 characters
            if (randomHex.Length < 8)
            {
                randomHex = randomHex.PadLeft(8, '0');
            }
            else if (randomHex.Length > 8)
            {
                randomHex = randomHex.Substring(0, 8);
            }

            return randomHex.ToUpper();
        }
    }
}