using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using D = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DP = DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using System.IO;
using System.IO.Compression;

namespace XMLyzeLibrary.Word
{
    public static class WF
    {
        public static (MainDocumentPart, Body) PopulateNewWordPackage(WordprocessingDocument package, UInt32Value? margin = null, string? border = null)
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
            // StyleDefinitionsPart stylePart = mainPart.AddNewPart<StyleDefinitionsPart>();
            // Styles styles = new(S.styleList);
            // styles.Save(stylePart);

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
    }
}