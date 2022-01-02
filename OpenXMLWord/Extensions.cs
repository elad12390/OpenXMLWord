using System;
using System.Collections.Generic;
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

        public static TResult Apply<T, TResult>(this T t, Func<T, TResult> fn) where T : class where TResult : class
            => t != null ? fn(t) : null;

        public static TResult Apply<T, TResult>(this T? t, Func<T, TResult> fn) where T : struct where TResult : class
            => t.HasValue ? fn(t.Value) : null;

        public static bool Apply<T>(this T t, Func<T, bool> fn) where T : class
            => t is {} && fn(t);

        public static T? Apply<T>(this bool val, T newVal) where T : struct
            => val ? (T?)newVal : null;

        public static T Apply<T>(this bool? val, T newVal) where T : class
            => val.HasValue && val.Value ? newVal : null;

        public static (MainDocumentPart mainPart, Body body, Header header, Footer footer) Initialize(this WordprocessingDocument doc, OpenXmlUtils.DocumentOptions options = null)
        {
            // Add a main document part. 
            var mainPart = doc.AddMainDocumentPart();

            mainPart.Document = new Document();
            
            var body = mainPart.Document.AppendChild(new Body());
            
            // Depends on body so need to come after creating the body
            Header header = null;
            if (options?.CreateHeader == true)
                header = IOpenXMLUtils.AddHeader(mainPart);
            
            Footer footer = null;
            if (options?.CreateFooter == true)
                footer = IOpenXMLUtils.AddFooter(mainPart);

            return (mainPart, body, header, footer);
        }
        
        public static (Table newTable, Table oldTable) CloneTableByTitle(this OpenXmlElement element, string title) =>
            OpenXmlUtils.CloneTableByTitle(element, title);

        public static Table FindTableByTitle(this OpenXmlElement element, string title) =>
            OpenXmlUtils.FindTableByTitle(element, title);

        public static void SetContentControl(this OpenXmlElement element, string tag, string value) =>
            OpenXmlUtils.SetContentControl(element, tag, value);

        public static void SetContentControls(this OpenXmlElement element, Dictionary<string, string> tagValueDictionary) =>
            OpenXmlUtils.SetContentControls(element, tagValueDictionary);

        public static void SetTableContentRows(this OpenXmlElement element, string tableTitle, List<Dictionary<string, string>> tableRows, int? maxRows = null) =>
            OpenXmlUtils.SetTableContentRows(element, tableTitle, tableRows, maxRows);

        public static Body Body(this WordprocessingDocument wordDocument) => wordDocument?.MainDocumentPart?.Document.Body;

        public static Header AddHeader(this MainDocumentPart mainDocumentPart) => OpenXmlUtils.AddHeader(mainDocumentPart);

        public static Footer AddFooter(this MainDocumentPart mainDocumentPart) => OpenXmlUtils.AddFooter(mainDocumentPart);
    }
}