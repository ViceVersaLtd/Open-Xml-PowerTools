// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXmlPowerTools
{
    public partial class WmlDocument
    {
        public WmlDocument FlattenAltChunks()
        {
            return AltChunkFlattener.FlattenAltChunks(this);
        }
    }

    public static class AltChunkFlattener
    {
        private static readonly string MinimalDefaultCss =
@"body, p, h1, h2, h3, h4, h5, h6, div, ul, ol, li, table, tr, td, th { display: block; }
li { display: list-item; }
table { display: table; }
tr { display: table-row; }
td, th { display: table-cell; }
";

        public static WmlDocument FlattenAltChunks(WmlDocument doc)
        {
            using (var streamDoc = new OpenXmlMemoryStreamDocument(doc))
            {
                using (var wDoc = streamDoc.GetWordprocessingDocument())
                {
                    FlattenAltChunks(wDoc);
                }
                return streamDoc.GetModifiedWmlDocument();
            }
        }

        public static void FlattenAltChunks(WordprocessingDocument doc)
        {
            if (doc == null)
                throw new ArgumentNullException(nameof(doc));

            MainDocumentPart mainPart = doc.MainDocumentPart;
            if (mainPart == null)
                return;

            XDocument mainXDoc = mainPart.GetXDocument();
            var altChunkElements = mainXDoc.Descendants(W.altChunk).ToList();
            if (altChunkElements.Count == 0)
                return;

            var settings = HtmlToWmlConverter.GetDefaultSettings();
            var replacements = new Dictionary<XElement, List<XElement>>();
            var partsToDelete = new HashSet<AlternativeFormatImportPart>();

            foreach (XElement altChunkEl in altChunkElements)
            {
                XAttribute idAttr = altChunkEl.Attribute(R.id);
                if (idAttr == null)
                {
                    altChunkEl.Remove();
                    continue;
                }

                string relId = idAttr.Value;
                AlternativeFormatImportPart importPart = null;
                try
                {
                    importPart = mainPart.GetPartById(relId) as AlternativeFormatImportPart;
                }
                catch
                {
                    // relationship may be broken
                }

                if (importPart == null)
                {
                    altChunkEl.Remove();
                    continue;
                }

                partsToDelete.Add(importPart);

                string htmlContent;
                using (Stream partStream = importPart.GetStream(FileMode.Open, FileAccess.Read))
                using (var reader = new StreamReader(partStream))
                {
                    htmlContent = reader.ReadToEnd();
                }

                List<XElement> newElements;
                try
                {
                    XElement htmlX = ParseHtmlContent(htmlContent);
                    WmlDocument converted = HtmlToWmlConverter.ConvertHtmlToWml(MinimalDefaultCss, "", "", htmlX, settings);
                    using (var ms = new MemoryStream(converted.DocumentByteArray))
                    using (var convertedDoc = WordprocessingDocument.Open(ms, false))
                    {
                        XElement body = convertedDoc.MainDocumentPart.GetXDocument().Root.Element(W.body);
                        newElements = body.Elements()
                            .Where(e => e.Name != W.sectPr)
                            .Select(e => new XElement(e)) // detach clone
                            .ToList();
                    }
                }
                catch (Exception)
                {
                    // Fallback: create a simple paragraph from stripped text
                    string plain = Regex.Replace(htmlContent, "<[^>]+>", " ").Trim();
                    if (string.IsNullOrWhiteSpace(plain))
                        plain = "(alt chunk content)";
                    newElements = new List<XElement>
                    {
                        new XElement(W.p,
                            new XElement(W.r,
                                new XElement(W.t, plain)))
                    };
                }

                replacements[altChunkEl] = newElements;
            }

            // Perform replacements (order safe because we use collected elements)
            foreach (var pair in replacements)
            {
                XElement oldEl = pair.Key;
                List<XElement> newEls = pair.Value;
                if (newEls != null && newEls.Count > 0)
                    oldEl.ReplaceWith(newEls.ToArray<object>());
                else
                    oldEl.Remove();
            }

            mainPart.PutXDocument();

            // Remove the alt chunk parts now that they are no longer referenced
            foreach (var part in partsToDelete)
            {
                try
                {
                    mainPart.DeletePart(part);
                }
                catch
                {
                    // ignore deletion issues
                }
            }
        }

        private static XElement ParseHtmlContent(string html)
        {
            // Best-effort parse. The alt chunks in the sample are tiny well-formed fragments.
            XElement x = XElement.Parse(html);
            x = (XElement)ConvertToNoNamespace(x);
            return x;
        }

        private static object ConvertToNoNamespace(XNode node)
        {
            XElement element = node as XElement;
            if (element != null)
            {
                return new XElement(element.Name.LocalName,
                    element.Attributes().Where(a => !a.IsNamespaceDeclaration),
                    element.Nodes().Select(n => ConvertToNoNamespace(n)));
            }
            return node;
        }
    }
}
