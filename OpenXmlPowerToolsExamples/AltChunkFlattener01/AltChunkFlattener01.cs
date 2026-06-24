/*******************************************************************************************
 * This example/test demonstrates flattening altChunks in a Word document.
 * 
 * Alt chunks (w:altChunk) are used to embed alternative content (commonly HTML) that Word
 * converts to native Word XML upon open/save. Some documents (e.g. certain exports or 
 * generated reports) contain alt chunks that need to be "flattened" into regular content
 * so that the document can be processed by DocumentBuilder, WmlComparer, etc.
 *
 * This test project:
 *  - Loads CAT A BLDG STANDARDS.docx (a document containing 13 altChunk HTML parts)
 *  - Calls WmlDocument.FlattenAltChunks() (or AltChunkFlattener.FlattenAltChunks)
 *  - Verifies that no altChunk elements remain
 *  - Saves the flattened result
 *
 * The implementation lives in OpenXmlPowerTools/AltChunkFlattener.cs and uses
 * HtmlToWmlConverter to turn the HTML chunks into native elements.
 *******************************************************************************************/

using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OpenXmlPowerTools;

class AltChunkFlattener01
{
    static void Main(string[] args)
    {
        string fixedResult = @"c:\Data\NZDevShed\ViceVersaLtd\Open-Xml-PowerTools\last-flattener-result.txt";
        try
        {
        var now = DateTime.Now;
        var outputDirName = string.Format(
            "ExampleOutput-{0:00}-{1:00}-{2:00}-{3:00}{4:00}{5:00}",
            now.Year - 2000, now.Month, now.Day, now.Hour, now.Minute, now.Second);
        var outputDir = new DirectoryInfo(outputDirName);
        outputDir.Create();

        Console.WriteLine("AltChunkFlattener01");
        Console.WriteLine("Output directory: " + outputDir.FullName);

        // Locate the sample. When running from the example folder the relative path works.
        // We also try a couple of common locations.
        string[] candidatePaths =
        {
            "NewDocxDocuments\\CatABldgStandards.docx",
            Path.Combine("..", "..", "..", "..", "NewDocxDocuments", "CatABldgStandards.docx"),
            @"..\..\..\NewDocxDocuments\CatABldgStandards.docx",
            @"..\..\..\..\NewDocxDocuments\CatABldgStandards.docx",
            Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "..", "NewDocxDocuments", "CatABldgStandards.docx"),
            @"c:\Data\NZDevShed\ViceVersaLtd\Open-Xml-PowerTools\NewDocxDocuments\CatABldgStandards.docx"
        };

        string inputPath = null;
        foreach (string cand in candidatePaths)
        {
            try
            {
                string full = Path.GetFullPath(cand);
                if (File.Exists(full))
                {
                    inputPath = full;
                    break;
                }
            }
            catch { }
        }

        if (inputPath == null)
        {
            // Hardcoded absolute for this dev environment / test run
            string hard = @"c:\Data\NZDevShed\ViceVersaLtd\Open-Xml-PowerTools\NewDocxDocuments\CatABldgStandards.docx";
            if (File.Exists(hard)) inputPath = hard;
        }

        if (inputPath == null)
        {
            Console.WriteLine("ERROR: Could not locate CatABldgStandards.docx");
            Console.WriteLine("Candidates tried:");
            foreach (var c in candidatePaths) Console.WriteLine("  " + c);
            Environment.Exit(1);
        }

        Console.WriteLine("Input document: " + inputPath);

        var sourceDoc = new WmlDocument(inputPath);
        int originalAltCount = CountAltChunks(sourceDoc);
        Console.WriteLine("AltChunk count before: " + originalAltCount);

        if (originalAltCount == 0)
        {
            Console.WriteLine("No alt chunks found. Nothing to do.");
            return;
        }

        WmlDocument flattened = sourceDoc.FlattenAltChunks();

        int afterAltCount = CountAltChunks(flattened);
        Console.WriteLine("AltChunk count after:  " + afterAltCount);

        string outputPath = Path.Combine(outputDir.FullName, "CatABldgStandards-Flattened.docx");
        flattened.SaveAs(outputPath);
        Console.WriteLine("Saved flattened document to: " + outputPath);

        if (afterAltCount == 0)
        {
            Console.WriteLine("SUCCESS: All alt chunks flattened.");
        }
        else
        {
            Console.WriteLine("WARNING: Some alt chunks remain.");
            Environment.Exit(2);
        }

        // Always write a machine readable result for verification when stdout capture is difficult
        var resultPath = Path.Combine(outputDir.FullName, "result.txt");
        File.WriteAllText(resultPath, string.Format("original={0}\nafter={1}\noutput={2}\n", originalAltCount, afterAltCount, outputPath));

        // Also write to a fixed location for easy tool inspection
        try { File.WriteAllText(@"c:\Data\NZDevShed\ViceVersaLtd\Open-Xml-PowerTools\last-flattener-result.txt", string.Format("original={0}\nafter={1}\noutput={2}\nsuccess={3}\n", originalAltCount, afterAltCount, outputPath, afterAltCount==0)); } catch {}

        // Also demonstrate the static method on an opened package
        using (var ms = new MemoryStream(flattened.DocumentByteArray))
        using (var pkgDoc = WordprocessingDocument.Open(ms, true))
        {
            AltChunkFlattener.FlattenAltChunks(pkgDoc);
            // (no-op here because already flattened, but shows the API)
        }
        }
        catch (Exception ex)
        {
            try { File.WriteAllText(fixedResult, "ERROR: " + ex.ToString()); } catch { }
            Console.WriteLine("EXCEPTION: " + ex.Message);
            Environment.Exit(99);
        }
    }

    private static int CountAltChunks(WmlDocument doc)
    {
        using (var ms = new MemoryStream(doc.DocumentByteArray))
        using (var wDoc = WordprocessingDocument.Open(ms, false))
        {
            return wDoc.MainDocumentPart.GetXDocument()
                .Descendants(W.altChunk)
                .Count();
        }
    }
}
