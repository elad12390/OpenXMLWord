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

        /// <summary>
        /// Set all content controls using the dictionary provided
        /// </summary>
        /// <param name="ancestor">the ancestor element to search in</param>
        /// <param name="tagValueDictionary">All tags(as key) and values(as value)</param>
        static void SetContentControls(OpenXmlElement ancestor, Dictionary<string, string> tagValueDictionary) =>
            OpenXMLWord.SetTextContentControls(ancestor, tagValueDictionary);

        /// <summary>
        /// Set picture on a picture control by tag
        /// </summary>
        /// <param name="doc">The word document</param>
        /// <param name="elm">The ancestor of the picture</param>
        /// <param name="tag">The tag of the picture control</param>
        /// <param name="imageType">Type of the image provided</param>
        /// <param name="fileStream">The image file</param>
        /// <exception cref="IOException">Will throw when given null stream</exception>
        static void SetContentControlImage(WordprocessingDocument doc, OpenXmlElement elm, string tag, ImagePartType imageType, FileStream fileStream) =>
            OpenXMLWord.SetContentControlImage(doc, elm, tag, imageType, fileStream);

        /// <summary>
        /// Set content control value
        /// </summary>
        /// <param name="elm">the parent element to search in</param>
        /// <param name="tag">the tag of the content control</param>
        /// <param name="value">the text value</param>
        static void SetTextContentControl(OpenXmlElement elm, string tag, string value) =>
            OpenXMLWord.SetTextContentControl(elm, tag, value);
        
        /// <summary>
        /// Clone a table by given title
        /// </summary>
        /// <param name="element">The parent element to search in his descendants</param>
        /// <param name="title">The table title</param>
        /// <returns>Tuple of the newTable, oldTable</returns>
        static (Table newTable, Table oldTable) CloneTableByTitle(OpenXmlElement element, string title) =>
            OpenXMLWord.CloneTableByTitle(element, title);

        /// <summary>
        /// Find a table by it's title
        /// </summary>
        /// <param name="element">Parent element to search in his descendants</param>
        /// <param name="title">The title of the table</param>
        /// <returns>The table element</returns>
        static Table FindTableByTitle(OpenXmlElement element, string title) =>
            OpenXMLWord.FindTableByTitle(element, title);

        /// <summary>
        /// Create a table with header, and rows multiplied by the contents provided
        /// if the table is overflowing will create another table with the same header
        /// </summary>
        /// <param name="element">The element to replace in</param>
        /// <param name="tableTitle">Title of the table to use</param>
        /// <param name="tableRows">The data for each row</param>
        /// <param name="maxRows">Max rows in a table</param>
        static void SetTableContentRows(OpenXmlElement element, string tableTitle, List<Dictionary<string, string>> tableRows, int? maxRows = null) =>
            OpenXMLWord.SetTableContentRows(element, tableTitle, tableRows, maxRows);

        /// <summary>
        /// Select the first matched element by their Tag (not all elements have this Tag)
        /// </summary>
        /// <param name="ancestor">the ancestor element to search in</param>
        /// <param name="tag">The tag to search for</param>
        /// <typeparam name="T">The type of the element searched for</typeparam>
        /// <returns>The element or null (if not found)</returns>
        public static T SelectElementByTag<T>(OpenXmlElement ancestor, string tag) where T : OpenXmlElement =>
            OpenXMLWord.SelectElementByTag<T>(ancestor, tag);
        
        /// <summary>
        /// Select elements by their Tag (not all elements have this Tag)
        /// </summary>
        /// <param name="ancestor">the ancestor element to search in</param>
        /// <param name="tag">The tag to search for</param>
        /// <typeparam name="T">The type of the element searched for</typeparam>
        /// <returns>list of matched elements or null (if not found)</returns>
        public static List<T> SelectElementsByTag<T>(OpenXmlElement ancestor, string tag) where T : OpenXmlElement =>
            OpenXMLWord.SelectElementsByTag<T>(ancestor, tag);
        
        /// <summary>
        /// Set all content controls using the dictionary provided
        /// </summary>
        /// <param name="element">The element to clone</param>
        /// <param name="content">The content to fill inside the content controls</param>
        /// <param name="createParagraphAfterClonedElement">Should create a paragraph between them or not</param>
        public static void CloneAndSetContent(OpenXmlElement element, List<Dictionary<string, string>> content, bool createParagraphAfterClonedElement = false) =>
            OpenXMLWord.CloneAndSetContent(element, content);
    }
}