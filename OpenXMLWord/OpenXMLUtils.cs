using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
namespace OpenXMLWord
{
    public class TableRowCreation : ITableRowCreation
    {
        public TableRowCreation(TableCreation parent, TableRow row)
        {
            Parent = parent;
            Row = row;
        }

        private TableCreation Parent { get; }
        private TableRow Row { get; }

        public ITableRowCreation CreateRowCell(OpenXmlCompositeElement cell, TableCellProperties properties)
        {
            // Create a cell.
            var tc = new TableCell();

            // Specify the width property of the table cell.
            tc.Append(properties);

            // Specify the table cell content.
            tc.Append(cell);

            // Append the table cell to the table row.
            Row.Append(tc);
        
            return this;
        }

        public ITableRowCreation CreateRowCells(IEnumerable<OpenXmlCompositeElement> cells, TableCellProperties properties)
        {
            foreach (var cell in cells)
                CreateRowCell(cell, properties);
            return this;
        }

        public ITableCreation EndRow() => Parent.AppendTableChild(Row);
    }

    public class TableCreation : ITableCreation
    {
        private OpenXmlCompositeElement Parent { get; }
        
        private Table Table { get; }
        
        public TableCreation(OpenXmlCompositeElement parent, TableProperties tblProp)
        {
            Parent = parent;
            Table = new Table();
            Table.AppendChild(tblProp);
        }

        public TableCreation AppendTableChild<T>(T newChild) where T : OpenXmlCompositeElement
        {
            Table.AppendChild(newChild);
            return this;
        }

        public ITableRowCreation CreateRow()
        {
            return new TableRowCreation(this, new TableRow());
        }

        public OpenXmlCompositeElement EndTable()
        {
            Parent.Append(Table);
            return Parent;
        }
    }

    public class ParagraphCreation : IParagraphCreation
    {
        private OpenXmlElement Parent { get; }
        
        private Paragraph Paragraph { get; }
        
        private ParagraphProperties ParagraphProperties { get {
                // If the paragraph has no ParagraphProperties object, create one.
                if (!Paragraph.Elements<ParagraphProperties>().Any())
                    Paragraph.PrependChild(new ParagraphProperties());
                
                // Get the paragraph properties element of the paragraph.
                return Paragraph.Elements<ParagraphProperties>().First();
        }}

        public ParagraphCreation(OpenXmlElement parent)
        {
            Parent = parent;
            Paragraph = new Paragraph();
        }
        
        public ParagraphCreation(OpenXmlElement parent, ParagraphStyleId styleId)
        {
            Parent = parent;
            Paragraph = new Paragraph(new ParagraphProperties(styleId));
        }

        public IParagraphCreation CreateText(string text)
        {
            var run = Paragraph.AppendChild(new Run());
            run.AppendChild(new Text(text));
            return this;
        }

        public IParagraphCreation AppendChild<T>(T newChild) where T : OpenXmlElement
        {
            Paragraph.AppendChild(newChild);
            return this;
        }

        public ParagraphCreation Align(JustificationValues justification)
        {
            ParagraphProperties.Justification ??= new Justification();
            ParagraphProperties.Justification.Val = justification;
            return this;
        }
        
        public IParagraphCreation ApplyStyle(WordprocessingDocument doc, OpenXmlUtils.StyleOptions styleOptions)
        {
            var pPr = ParagraphProperties;
            
            // Get the Styles part for this document.
            var part = doc.MainDocumentPart?.StyleDefinitionsPart;
            
            // If the Styles part does not exist, add it and then add the style.
            if (part == null)
            {
                part = OpenXmlUtils.AddStylesPartToPackage(doc);
                OpenXmlUtils.AddNewStyle(part, styleOptions);
            }
            else
            {
                // If the style is not in the document, add it.
                if (OpenXmlUtils.IsStyleIdInDocument(doc, styleOptions.StyleId) != true)
                {
                    // No match on styleId, so let's try style name.
                    var styleIdFromName = OpenXmlUtils.GetStyleIdFromStyleName(doc, styleOptions.StyleName, StyleValues.Paragraph);
                    if (styleIdFromName == null)
                        OpenXmlUtils.AddNewStyle(part, styleOptions);
                    else
                        styleOptions.StyleId = styleIdFromName;
                }
            }

            // Set the style of the paragraph.
            pPr.ParagraphStyleId = new ParagraphStyleId { Val = styleOptions.StyleId };
            return this;
        }

        public OpenXmlElement EndParagraph()
        {
            Parent.Append(Paragraph);
            return Parent;
        }
    }

    public class OpenXmlUtils : IOpenXMLUtils
    {
        private static Int64Value emusPerCm = 360000U;
        private static UInt32Value _imageId = 0U;

        public class DocumentOptions
        {
            public bool CreateHeader { get; set; }
            public bool CreateFooter { get; set; }
        }
        public class ImageOptions
        {
            public class Transform
            {
                /// <summary>
                /// Size in CM
                /// </summary>
                public float? SizeX { get; set; }
                public Int64Value SizeXPerCm => SizeX.Apply(size => (Int64Value)(size * emusPerCm));
                /// <summary>
                /// Height in CM
                /// </summary>
                public float? SizeY { get; set; }
                public Int64Value SizeYPerCm => SizeY.Apply(size => (Int64Value)(size * emusPerCm));
                /// <summary>
                /// Offset in CM
                /// </summary>
                public float? OffsetX { get; set; }
                public Int64Value OffsetXPerCm => OffsetX.Apply(size => (Int64Value)(size * emusPerCm));
                /// <summary>
                /// Offset in CM
                /// </summary>
                public float? OffsetY { get; set; }
                public Int64Value OffsetYPerCm => OffsetY.Apply(size => (Int64Value)(size * emusPerCm));
            }

            public class Crop
            {
                public Int64Value LeftEdge { get; set; }
                public Int64Value TopEdge { get; set; }
                public Int64Value RightEdge { get; set; }
                public Int64Value BottomEdge { get; set; }
            }

            public class Margin
            {
                public UInt32Value Top { get; set; }
                public UInt32Value Bottom { get; set; }
                public UInt32Value Left { get; set; }
                public UInt32Value Right { get; set; }
            }
            public Transform Trans { get; set; }
            public Margin Marg { get; set; }
            public Crop Cro { get; set; }
            public string Name { get; set; }
            public bool? NoChangeAspect { get; set; }
        }

        public class StyleOptions
        {
            public string StyleId { get; set; }
            public string StyleName { get; set; }
            public string BasedOn { get; set; }
            public ThemeColorValues? Color { get; set; }
            public string Font { get; set; }
            /// <summary>
            /// Font size in pt
            /// </summary>
            public double? FontSize { get; set; }
            public string FontSizeVal => (FontSize * 2).ToString();
            public bool? Bold { get; set; }
            public bool? Italic { get; set; }
        }

        public static Header AddHeader(MainDocumentPart mainDocumentPart)
        {
            //Delete the existing header parts.
            mainDocumentPart.DeleteParts(mainDocumentPart.HeaderParts);

            //Create a new header part and get its relationship id.
            var newHeaderPart = mainDocumentPart.AddNewPart<HeaderPart>();

            // Create the header in the new header part
            newHeaderPart.Header = new Header();

            if (!mainDocumentPart.Document.Descendants<SectionProperties>().Any())
            {
                mainDocumentPart.Document.Body?.AppendChild(new SectionProperties(
                    new HeaderReference
                    {
                        Id = mainDocumentPart.GetIdOfPart(newHeaderPart),
                        Type = HeaderFooterValues.Default
                    }
                ));
            }
            else
            {
                //Loop through all section properties in the document
                //which is where header references are defined.
                foreach (var sectProperties in mainDocumentPart.Document.Descendants<SectionProperties>())
                {
                    // Delete any existing references to headers.
                    foreach (var headerReference in sectProperties.Descendants<HeaderReference>())
                        sectProperties.RemoveChild(headerReference);

                    // Create a new header reference that points to the new
                    //header part and add it to the section properties.
                    var newHeaderReference = new HeaderReference { Id = mainDocumentPart.GetIdOfPart(newHeaderPart), Type = HeaderFooterValues.Default };
                    sectProperties.Append(newHeaderReference);
                }
            }

            return newHeaderPart.Header; 
        }

        public static Footer AddFooter(MainDocumentPart mainDocumentPart)
        {
            //Delete the existing header parts.
            mainDocumentPart.DeleteParts(mainDocumentPart.FooterParts);

            //Create a new header part and get its relationship id.
            var newFooterPart = mainDocumentPart.AddNewPart<FooterPart>();

            // Create the header in the new header part
            newFooterPart.Footer = new Footer();

            //Loop through all section properties in the document
            //which is where header references are defined.
            foreach (var sectProperties in mainDocumentPart.Document.Descendants<SectionProperties>())
            {
                // Delete any existing references to headers.
                foreach (var footerReference in sectProperties.Descendants<FooterReference>())
                    sectProperties.RemoveChild(footerReference);

                // Create a new header reference that points to the new
                //header part and add it to the section properties.
                var newFooterReference = new FooterReference() { Id = mainDocumentPart.GetIdOfPart(newFooterPart), Type = HeaderFooterValues.Default };
                sectProperties.Append(newFooterReference);
            }

            return newFooterPart.Footer;
        }

        
        public static Run CreateImage(MainDocumentPart mainPart, string imageUrl, ImagePartType imageType, ImageOptions ops = null)
        {
            var imagePart = mainPart.AddImagePart(imageType);

            using (var stream = new FileStream(imageUrl, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            var imageId = _imageId++;
            
            var element =
                new Drawing(
                    new DW.Inline(
                        new DW.Extent { Cx = ops?.Trans?.SizeXPerCm, Cy = ops?.Trans?.SizeYPerCm },
                        new DW.EffectExtent
                        {
                            LeftEdge = ops?.Cro?.LeftEdge ?? 0,
                            TopEdge = ops?.Cro?.TopEdge ?? 0,
                            RightEdge = ops?.Cro?.RightEdge ?? 0,
                            BottomEdge = ops?.Cro?.BottomEdge ?? 0
                        },
                        new DW.DocProperties()
                        {
                            Id = imageId,
                            Name = ops?.Name ?? "Image"
                        },
                        new DW.NonVisualGraphicFrameDrawingProperties(new A.GraphicFrameLocks() { NoChangeAspect = ops?.NoChangeAspect }),
                        new A.Graphic(
                            new A.GraphicData(
                                new PIC.Picture(
                                    new PIC.NonVisualPictureProperties(
                                        new PIC.NonVisualDrawingProperties
                                        {
                                            Id = imageId,
                                            Name = ops?.Name ?? "New Image.jpg"
                                        },
                                        new PIC.NonVisualPictureDrawingProperties()
                                    ),
                                    new PIC.BlipFill(
                                        new A.Blip(new A.BlipExtensionList(new A.BlipExtension { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" }))
                                        {
                                            Embed = mainPart.GetIdOfPart(imagePart),
                                            CompressionState = A.BlipCompressionValues.Print
                                        },
                                        new A.Stretch(new A.FillRectangle())
                                    ),
                                    new PIC.ShapeProperties(
                                        new A.Transform2D(
                                            new A.Offset { X = ops?.Trans?.OffsetXPerCm, Y = ops?.Trans?.OffsetYPerCm },
                                            new A.Extents { Cx =  ops?.Trans?.SizeXPerCm, Cy = ops?.Trans?.SizeYPerCm }
                                        ),
                                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                                    )
                                )
                            ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
                        )
                    )
                    {
                        DistanceFromTop = ops?.Marg?.Top,
                        DistanceFromBottom = ops?.Marg?.Bottom,
                        DistanceFromLeft = ops?.Marg?.Left,
                        DistanceFromRight = ops?.Marg?.Right,
                        EditId = "50D07946"
                    }
                );
            
            return new Run(element);
        }
        
        public static StyleDefinitionsPart AddStylesPartToPackage(WordprocessingDocument doc)
        {
            StyleDefinitionsPart part;
            part = doc.MainDocumentPart?.AddNewPart<StyleDefinitionsPart>();
            var root = new Styles();
            root.Save(part);
            return part;
        }
        
        // Create a new style with the specified styleId and styleName and add it to the specified
        // style definitions part.
        public static void AddNewStyle(StyleDefinitionsPart styleDefinitionsPart, StyleOptions ops)
        {
            // Get access to the root element of the styles part.
            var styles = styleDefinitionsPart.Styles;

            // Create a new paragraph style and specify some of the properties.
            var style = new Style { Type = StyleValues.Paragraph, StyleId = ops.StyleId, CustomStyle = true };
            var styleName = new StyleName { Val = ops.StyleName };
            var basedOn = new BasedOn { Val = ops.BasedOn };
            var nextParagraphStyle = new NextParagraphStyle() { Val = "Normal" };
            
            style.Append(styleName);
            style.Append(basedOn);
            style.Append(nextParagraphStyle);

            // Create the StyleRunProperties object and specify some of the run properties.
            var styleRunProperties = new StyleRunProperties()
            {
                Color = ops.Color.Apply(c => new Color { ThemeColor = c}),
                RunFonts = ops.Font.Apply(f => new RunFonts {Ascii = f}),
                FontSize = ops.FontSizeVal.Apply(f => new FontSize { Val = f}),
                Bold = ops.Bold.Apply(new Bold()),
                Italic = ops.Italic.Apply(new Italic()),
            };
        
            // Add the run properties to the style.
            style.Append(styleRunProperties);
        
            // Add the style to the styles part.
            styles?.Append(style);
        }
        
        // Return true if the style id is in the document, false otherwise.
        public static bool IsStyleIdInDocument(WordprocessingDocument doc, string styleId)
        {
            // Get access to the Styles element for this document.
            var s = doc?.MainDocumentPart?.StyleDefinitionsPart?.Styles;

                    // Check that styles exists
            return s is { }
                   // Check that the style exists
                   && s.Elements<Style>().Any(st => st.StyleId == styleId && st.Type == StyleValues.Paragraph);
        }

        // Return styleId that matches the styleName, or null when there's no match.
        public static string GetStyleIdFromStyleName(WordprocessingDocument doc, string styleName, StyleValues styleValue)
        {
            var styles = doc?.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            string styleId = styles?
                .Descendants<StyleName>()
                .Where(s => s.Val.Apply(val => styleName.Equals(val)) && ((Style)s.Parent)?.Type == styleValue)
                .Select(n => ((Style)n.Parent)?.StyleId)
                .FirstOrDefault();
            return styleId;
        }

        // Set all content controls using the dictionary provided
        public static void SetContentControls(OpenXmlElement elm, Dictionary<string, string> tagValueDictionary)
        {
            foreach (var (tag, value) in tagValueDictionary)
                SetContentControl(elm, tag, value);
        }

        // Set the content control value by tag
        public static void SetContentControl(OpenXmlElement elm, string tag, string value)
        {
            var elements = elm.Descendants<SdtElement>()
                .Where(elm => elm.SdtProperties?.GetFirstChild<Tag>()?.Val == tag)
                .ToList();

            var descendants = elements.SelectMany(element => element.Descendants<Text>()).Where(t => t is { });
            foreach (var descendant in descendants)
                descendant.Text = value;
        }

        // Set all content controls using the dictionary provided
        public static void SetTableContentRows(OpenXmlElement element, string tableTitle, List<Dictionary<string, string>> tableRows, int? maxRows = null)
        {
            var table = element.FindTableByTitle(tableTitle);
            var lastRow = table.Descendants<TableRow>().ElementAtOrDefault(1);
            lastRow?.Remove();
            
            for (var i = 0; i < tableRows.Count; i++)
            {
                var row = tableRows[i];
                var index = i;
                var newRow = (TableRow)lastRow.CloneNode(true);
                SetContentControls(newRow, row);
                table.AppendChild(newRow);
                if (maxRows == null || ((i + 1) % maxRows != 0) || (i + 1) >= tableRows.Count) continue;
                
                // Create new table (max rows reached)
                var (newTable, _) = CloneTableByTitle(element, tableTitle);
                foreach (var tableRow in newTable.Descendants<TableRow>().Skip(1).ToList())
                    tableRow.Remove();
                var p = table.Parent?.InsertAfter(new Paragraph(), table);
                table.Parent?.InsertAfter(newTable, p);
                table = newTable;
            }
        }
        
        // Clone table
        public static (Table newTable, Table oldTable) CloneTableByTitle(OpenXmlElement element, string title)
        {
            var table = element.Descendants<Table>()
                .FirstOrDefault(table => table.Descendants<TableCaption>().FirstOrDefault()?.Val == title);

            return ((Table)table?.CloneNode(true), table);
        }
        
        // Clone table
        public static Table FindTableByTitle(OpenXmlElement element, string title)
        {
            var table = element
                .Descendants<Table>()
                .FirstOrDefault(table => table.Descendants<TableCaption>().FirstOrDefault()?.Val == title);

            return table;
        }
        
    }
}