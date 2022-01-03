using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLWord
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            // CreateDocumentManuallyExample();
            CreateDocumentFromTemplateExample();
        }

        private static void CreateDocumentFromTemplateExample()
        {
            // Create a document by supplying the filepath. 
            using var wordDocument = WordprocessingDocument.CreateFromTemplate("D:/Work/Wanzy/Template1WithFields.dotx");
            var tableRows = new List<Dictionary<string, string>>();
            tableRows.AddRange(Enumerable.Repeat(new Dictionary<string, string>()
            {
                { "FieldName", string.Join(' ', GetNRandomStrings(3)) },
                { "FieldValue", string.Join(' ', GetNRandomStrings(10)) },
            }, 100).ToList());
            // wordDocument.Body().SetTableContentRows("test", tableRows, maxRows: 2);
            // IOpenXMLWord.SetContentControls(wordDocument.Body(), new Dictionary<string, string>
            // {
            //     {"PersonalWords", string.Join(' ', GetNRandomStrings(25))},
            //     {"StudentName", "Elad meow"}
            // });
            IOpenXMLWord.SetContentControls(wordDocument.Body(), new Dictionary<string, string>
            {
                {"StudentName", "Elad"},
            });
            
            using (var f = File.OpenRead("Path/To/File"))
                wordDocument.Body().SetContentControlImage(wordDocument, tag: "Pic1", imageType:ImagePartType.Png, fileStream: f);
            
            // using (var f = File.OpenRead("Path/To/File"))
            //     wordDocument.Body().SetContentControlImage(wordDocument, tag: "Pic2", imageType:ImagePartType.Png, fileStream: f);
            wordDocument.SaveAs("./Test.docx");
            wordDocument.Close();
        }

        private static void CreateDocumentManuallyExample()
        {
            // Create a document by supplying the filepath. 
            using var wordDocument = WordprocessingDocument.Create("word.doc", WordprocessingDocumentType.Document);

            // Initialize MainPart and Body
            var (mainPart, body, _, _) = wordDocument.Initialize();

            var table = body
                .CreateTable(
                    new TableProperties(
                        new TableBorders(
                            new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 24 },
                            new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 24 },
                            new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 24 },
                            new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 24 },
                            new InsideHorizontalBorder
                                { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 24 },
                            new InsideVerticalBorder
                                { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 24 }
                        )
                    )
                );

            foreach (var _ in new[] { 1, 2, 3, 4, 5, 6, 7 })
            {
                table.CreateRow()
                    .CreateRowCell(
                        new Paragraph(new Run(new Text("some text"))),
                        new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2400" })
                    )
                    .CreateRowCell(
                        new Paragraph(new Run(new Text("some text 2"))),
                        new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = "2400" })
                    )
                    .EndRow();
            }

            table.EndTable();

            body
                .CreateParagraph()
                .Align(JustificationValues.Center)
                .ApplyStyle(wordDocument, new StyleOptions
                {
                    StyleId = "cool_new_style",
                    StyleName = "Cool Style",
                    Bold = true,
                    Color = ThemeColorValues.Accent2,
                    Font = "Arial",
                    Italic = true,
                    BasedOn = "Normal",
                    FontSize = 24,
                })
                .CreateText("Cool Text")
                .CreateText("Cool Text 2")
                .CreateText("Cool Text 3")
                .CreateImage(
                    mainPart,
                    "C:/Users/elad1/Downloads/carbon (15).png",
                    ImagePartType.Png,
                    new ImageOptions()
                    {
                        Trans = new ImageOptions.Transform()
                        {
                            SizeX = .5f,
                            SizeY = .5f
                        }
                    }
                )
                .EndParagraph();
        }

        private static IEnumerable<string> GetNRandomStrings(int n)
        {
            var wordRandomizer = new EnglishWordRandomizer();
            for (var i = 0; i < n; i++)
                yield return wordRandomizer.Next();
        }

        private class EnglishWordRandomizer
        {
            private Random Random { set; get; }
            private EnglishCharRandomizer EnglishCharRandomizer { get; }
            private int MaxWordLength { get; set; }

            public EnglishWordRandomizer()
            {
                EnglishCharRandomizer = new EnglishCharRandomizer();
                MaxWordLength = 20;
                Random = new Random();
            }

            public string Next()
            {
                var wordLength = Random.Next(1, MaxWordLength);
                var word = new StringBuilder(MaxWordLength);
                for (var i = 0; i < wordLength; i++)
                {
                    var c = EnglishCharRandomizer.Next();
                    word.Append(c);
                }

                return word.ToString();
            }
        }

        private class EnglishCharRandomizer
        {
            private Random Random { set; get; }

            public EnglishCharRandomizer()
            {
                Random = new Random();
            }

            public char Next()
            {
                var index = Random.Next(0, 'z' - 'a');
                var isUpper = Random.Next(0, 101) > 80;
                return (char)((isUpper ? 'A' : 'a') + index);
            }
        } 
    }
}