using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
namespace OpenXMLWord
{
    public interface ITableRowCreation
    {
        public ITableRowCreation CreateRowCell(OpenXmlCompositeElement cell, TableCellProperties properties);

        public ITableRowCreation CreateRowCells(IEnumerable<OpenXmlCompositeElement> cells, TableCellProperties properties);
        
        public ITableCreation EndRow();
    }

    public interface ITableCreation
    {
        public ITableRowCreation CreateRow();

        public OpenXmlCompositeElement EndTable();
    }

    public interface IParagraphCreation
    {
        IParagraphCreation AppendChild<T>(T newChild) where T : OpenXmlElement;
        
        IParagraphCreation CreateImage(MainDocumentPart mainPart, string imageUrl, ImagePartType imageType, ImageOptions ops = null)
            => AppendChild(IOpenXMLWord.CreateImage(mainPart, imageUrl, imageType, ops));
        
        OpenXmlElement EndParagraph();
        
        IParagraphCreation ApplyStyle(WordprocessingDocument wordDocument, StyleOptions styleOptions);
        
        IParagraphCreation CreateText(string text);
        
        ParagraphCreation Align(JustificationValues justification);
    }

    public interface IHeaderCreation
    {
        
    }

    public interface IOpenXMLWord
    {
        public static Run CreateImage(MainDocumentPart mainPart, string imageUrl, ImagePartType imageType, ImageOptions ops = null) =>
            OpenXMLWord.CreateImage(mainPart, imageUrl, imageType, ops);

        static Header AddHeader(MainDocumentPart mainDocumentPart) => OpenXMLWord.AddHeader(mainDocumentPart);

        static Footer AddFooter(MainDocumentPart mainDocumentPart) => OpenXMLWord.AddFooter(mainDocumentPart);

        static void SetContentControls(OpenXmlElement element, Dictionary<string, string> tagValueDictionary) =>
            OpenXMLWord.SetContentControls(element, tagValueDictionary);

        static void SetContentControlImage(WordprocessingDocument doc, OpenXmlElement elm, string tag, ImagePartType imageType, FileStream fileStream) =>
            OpenXMLWord.SetContentControlImage(doc, elm, tag, imageType, fileStream);

        static void SetContentControl(OpenXmlElement elm, string tag, string value) =>
            OpenXMLWord.SetContentControl(elm, tag, value);

        static (Table newTable, Table oldTable) CloneTableByTitle(OpenXmlElement element, string title) =>
            OpenXMLWord.CloneTableByTitle(element, title);

        static Table FindTableByTitle(OpenXmlElement element, string title) =>
            OpenXMLWord.FindTableByTitle(element, title);

        static void SetTableContentRows(OpenXmlElement element, string tableTitle, List<Dictionary<string, string>> tableRows, int? maxRows = null) =>
            OpenXMLWord.SetTableContentRows(element, tableTitle, tableRows, maxRows);
    }
}