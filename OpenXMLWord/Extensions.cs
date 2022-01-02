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
        
        public static (Table newTable, Table oldTable) CloneTableByTitle(this OpenXmlElement element, string title) =>
            IOpenXMLWord.CloneTableByTitle(element, title);

        public static Table FindTableByTitle(this OpenXmlElement element, string title) =>
            IOpenXMLWord.FindTableByTitle(element, title);

        public static void SetContentControl(this OpenXmlElement element, string tag, string value) =>
            IOpenXMLWord.SetContentControl(element, tag, value);

        public static void SetContentControls(this OpenXmlElement element, Dictionary<string, string> tagValueDictionary) =>
            IOpenXMLWord.SetContentControls(element, tagValueDictionary);

        public static void SetTableContentRows(this OpenXmlElement element, string tableTitle, List<Dictionary<string, string>> tableRows, int? maxRows = null) =>
            IOpenXMLWord.SetTableContentRows(element, tableTitle, tableRows, maxRows);

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
    }
}