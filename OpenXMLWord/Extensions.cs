using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OpenXMLWord
{
    public static class Extensions
    {
        public static ITableCreation CreateTable(this OpenXmlCompositeElement element, TableProperties tblProp)
            => new TableCreation(element, tblProp);
        
        public static IParagraphCreation CreateParagraph(this OpenXmlElement element)
            => new ParagraphCreation(element);
        
        public static IParagraphCreation CreateParagraph(this OpenXmlElement element, ParagraphStyleId styleId)
            => new ParagraphCreation(element, styleId);

        public static (MainDocumentPart mainPart, Body body, Header header, Footer footer) Initialize(this WordprocessingDocument doc, DocumentOptions options = null)
        {
            // Add a main document part. 
            var mainPart = doc.AddMainDocumentPart();

            mainPart.Document = new Document();
            
            var body = mainPart.Document.AppendChild(new Body());
            
            // Depends on body so need to come after creating the body
            Header header = null;
            if (options?.CreateHeader == true)
                header = IOpenXMLWord.AddHeader(mainPart);
            
            Footer footer = null;
            if (options?.CreateFooter == true)
                footer = IOpenXMLWord.AddFooter(mainPart);

            return (mainPart, body, header, footer);
        }
        
        /// <summary>
        /// Clone a table by given title
        /// </summary>
        /// <param name="element">The parent element to search in his descendants</param>
        /// <param name="title">The table title</param>
        /// <returns>Tuple of the newTable, oldTable</returns>
        public static (Table newTable, Table oldTable) CloneTableByTitle(this OpenXmlElement element, string title) =>
            IOpenXMLWord.CloneTableByTitle(element, title);

        /// <summary>
        /// Find a table by it's title
        /// </summary>
        /// <param name="element">Parent element to search in his descendants</param>
        /// <param name="title">The title of the table</param>
        /// <returns>The table element</returns>
        public static Table FindTableByTitle(this OpenXmlElement element, string title) =>
            IOpenXMLWord.FindTableByTitle(element, title);

        /// <summary>
        /// Set content control value
        /// </summary>
        /// <param name="elm">the parent element to search in</param>
        /// <param name="tag">the tag of the content control</param>
        /// <param name="value">the text value</param>
        public static void SetContentControl(this OpenXmlElement element, string tag, string value) =>
            IOpenXMLWord.SetTextContentControl(element, tag, value);

        /// <summary>
        /// Set all content controls using the dictionary provided
        /// </summary>
        /// <param name="ancestor">the ancestor element to search in</param>
        /// <param name="tagValueDictionary">All tags(as key) and values(as value)</param>
        public static void SetContentControls(this OpenXmlElement ancestor, Dictionary<string, string> tagValueDictionary) =>
            IOpenXMLWord.SetContentControls(ancestor, tagValueDictionary);

        /// <summary>
        /// Create a table with header, and rows multiplied by the contents provided
        /// if the table is overflowing will create another table with the same header
        /// </summary>
        /// <param name="element">The element to replace in</param>
        /// <param name="tableTitle">Title of the table to use</param>
        /// <param name="tableRows">The data for each row</param>
        /// <param name="maxRows">Max rows in a table</param>
        public static void SetTableContentRows(this OpenXmlElement element, string tableTitle, List<Dictionary<string, string>> tableRows, int? maxRows = null) =>
            IOpenXMLWord.SetTableContentRows(element, tableTitle, tableRows, maxRows);

        /// <summary>
        /// Set picture on a picture control by tag
        /// </summary>
        /// <param name="doc">The word document</param>
        /// <param name="elm">The ancestor of the picture</param>
        /// <param name="tag">The tag of the picture control</param>
        /// <param name="imageType">Type of the image provided</param>
        /// <param name="fileStream">The image file</param>
        /// <exception cref="IOException">Will throw when given null stream</exception>
        public static void SetContentControlImage(this OpenXmlElement elm, WordprocessingDocument doc, string tag, ImagePartType imageType, FileStream fileStream) =>
            IOpenXMLWord.SetContentControlImage(doc, elm, tag, imageType, fileStream);

        public static Body Body(this WordprocessingDocument wordDocument) => wordDocument?.MainDocumentPart?.Document.Body;

        public static Header AddHeader(this MainDocumentPart mainDocumentPart) => OpenXMLWord.AddHeader(mainDocumentPart);

        public static Footer AddFooter(this MainDocumentPart mainDocumentPart) => OpenXMLWord.AddFooter(mainDocumentPart);

        internal static TResult Apply<T, TResult>(this T t, Func<T, TResult> fn) where T : class where TResult : class
            => t != null ? fn(t) : null;

        internal static TResult Apply<T, TResult>(this T? t, Func<T, TResult> fn) where T : struct where TResult : class
            => t.HasValue ? fn(t.Value) : null;

        internal static bool Apply<T>(this T t, Func<T, bool> fn) where T : class
            => t is {} && fn(t);

        internal static T? Apply<T>(this bool val, T newVal) where T : struct
            => val ? (T?)newVal : null;

        internal static T Apply<T>(this bool? val, T newVal) where T : class
            => val.HasValue && val.Value ? newVal : null;
        
        internal static void ForEach<T>(this IEnumerable<T> enumeration, Action<T> action)
        {
            foreach(var item in enumeration)
                action(item);
        }

        /// <summary>
        /// Select the first matched element by their Tag (not all elements have this Tag)
        /// </summary>
        /// <param name="ancestor">the ancestor element to search in</param>
        /// <param name="tag">The tag to search for</param>
        /// <typeparam name="T">The type of the element searched for</typeparam>
        /// <returns>The element or null (if not found)</returns>
        public static T SelectElementByTag<T>(this OpenXmlElement ancestor, string tag) where T : OpenXmlElement =>
            IOpenXMLWord.SelectElementByTag<T>(ancestor, tag);
        
        /// <summary>
        /// Select elements by their Tag (not all elements have this Tag)
        /// </summary>
        /// <param name="ancestor">the ancestor element to search in</param>
        /// <param name="tag">The tag to search for</param>
        /// <typeparam name="T">The type of the element searched for</typeparam>
        /// <returns>list of matched elements or null (if not found)</returns>
        public static List<T> SelectElementsByTag<T>(this OpenXmlElement ancestor, string tag) where T : OpenXmlElement =>
            IOpenXMLWord.SelectElementsByTag<T>(ancestor, tag);
        
        /// <summary>
        /// Set all content controls using the dictionary provided
        /// </summary>
        /// <param name="element">The element to clone</param>
        /// <param name="content">The content to fill inside the content controls</param>
        /// <param name="createParagraphAfterClonedElement">Should create a paragraph between them or not</param>
        public static void CloneAndSetContent(this OpenXmlElement element, List<Dictionary<string, string>> content, bool createParagraphAfterClonedElement = false) =>
            IOpenXMLWord.CloneAndSetContent(element, content);
    }
}