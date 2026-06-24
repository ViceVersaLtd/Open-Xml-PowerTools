// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.IO;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;
using Xunit;

#if !ELIDE_XUNIT_TESTS

namespace OxPt
{
    public class AltChunkFlattenerTests
    {
        [Fact]
        public void AC001_FlattenCatABldgStandards()
        {
            // The sample is copied into NewDocxDocuments at solution root during dev setup.
            string sample = Path.GetFullPath(Path.Combine("..", "..", "..", "..", "NewDocxDocuments", "CatABldgStandards.docx"));
            if (!File.Exists(sample))
            {
                // fallback for when tests run from bin dirs etc.
                sample = @"c:\Data\NZDevShed\ViceVersaLtd\Open-Xml-PowerTools\NewDocxDocuments\CatABldgStandards.docx";
            }
            Assert.True(File.Exists(sample), "Sample alt-chunk document not found at: " + sample);

            var input = new WmlDocument(sample);
            int before = CountAltChunks(input);
            Assert.True(before > 0, "Expected alt chunks in the sample document");

            var output = input.FlattenAltChunks();

            int after = CountAltChunks(output);
            Assert.Equal(0, after);

            // Also exercise the direct on WordprocessingDocument overload
            using (var ms = new MemoryStream(output.DocumentByteArray))
            using (var wDoc = WordprocessingDocument.Open(ms, true))
            {
                AltChunkFlattener.FlattenAltChunks(wDoc);
            }
        }

        private static int CountAltChunks(WmlDocument doc)
        {
            using (var ms = new MemoryStream(doc.DocumentByteArray))
            using (var wDoc = WordprocessingDocument.Open(ms, false))
            {
                return wDoc.MainDocumentPart.GetXDocument().Descendants(W.altChunk).Count();
            }
        }
    }
}

#endif
