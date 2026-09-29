// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using Xunit;

using static DocumentFormat.OpenXml.Tests.TestAssets;

using P = DocumentFormat.OpenXml.Presentation;
using S = DocumentFormat.OpenXml.Spreadsheet;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentFormat.OpenXml.Tests
{
    public class XmlNamespacePrefixSettingsTests
    {
        private const string XmlDeclaration = @"<?xml version=""1.0"" encoding=""utf-8""?>";
        private const string SpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string WordNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        private const string StrictWordNs = "http://purl.oclc.org/ooxml/wordprocessingml/main";
        private const string PresentationNs = "http://schemas.openxmlformats.org/presentationml/2006/main";
        private const string RelationshipsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        [Fact]
        public void PresetWritesWorksheetInDefaultNamespace()
        {
            using var stream = new MemoryStream();
            using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook).UseDefaultNamespaceForRoot();
            var worksheetPart = AddWorksheet(doc, new Worksheet(new SheetData(new Row(new Cell(new CellValue("1"))))));

            Assert.Equal(
                $@"{XmlDeclaration}<worksheet xmlns=""{SpreadsheetNs}""><sheetData><row><c><v>1</v></c></row></sheetData></worksheet>",
                SaveAndRead(worksheetPart));
        }

        [Fact]
        public void PresetDeclaresBuiltInPrefixForAttributesInDefaultNamespace()
        {
            using var stream = new MemoryStream();
            using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document).UseDefaultNamespaceForRoot();
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new W.Document(new W.Body(new W.Paragraph(new W.ParagraphProperties(new W.Justification { Val = W.JustificationValues.Center }))));

            Assert.Equal(
                $@"{XmlDeclaration}<document xmlns:w=""{WordNs}"" xmlns=""{WordNs}""><body><p><pPr><jc w:val=""center"" /></pPr></p></body></document>",
                SaveAndRead(mainPart));
        }

        [Fact]
        public void PresetWritesPresentationInDefaultNamespace()
        {
            using var stream = new MemoryStream();
            using var doc = PresentationDocument.Create(stream, PresentationDocumentType.Presentation).UseDefaultNamespaceForRoot();
            var presentationPart = doc.AddPresentationPart();
            presentationPart.Presentation = new P.Presentation(new P.SlideIdList());

            Assert.Equal(
                $@"{XmlDeclaration}<presentation xmlns=""{PresentationNs}""><sldIdLst /></presentation>",
                SaveAndRead(presentationPart));
        }

        [Fact]
        public void WithoutSettingsBuiltInPrefixIsUsed()
        {
            using var stream = new MemoryStream();
            using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            var worksheetPart = AddWorksheet(doc, new Worksheet(new SheetData()));

            Assert.Equal($@"{XmlDeclaration}<x:worksheet xmlns:x=""{SpreadsheetNs}""><x:sheetData /></x:worksheet>", SaveAndRead(worksheetPart));
        }

        [Fact]
        public void DictionaryEntryMakesRootNamespaceDefault()
        {
            var settings = new XmlNamespacePrefixSettings();
            settings.Prefixes[SpreadsheetNs] = string.Empty;

            using var stream = new MemoryStream();
            using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook).UseNamespacePrefixes(settings);
            var worksheetPart = AddWorksheet(doc, new Worksheet(new SheetData()));

            Assert.Equal($@"{XmlDeclaration}<worksheet xmlns=""{SpreadsheetNs}""><sheetData /></worksheet>", SaveAndRead(worksheetPart));
        }

        [Fact]
        public void OnlyRootNamespaceBecomesDefault()
        {
            var settings = new XmlNamespacePrefixSettings();
            settings.Prefixes[SpreadsheetNs] = string.Empty;
            settings.Prefixes[RelationshipsNs] = string.Empty;

            using var stream = new MemoryStream();
            using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook).UseNamespacePrefixes(settings);
            var worksheetPart = AddWorksheet(doc, new Worksheet(new SheetData(), new S.Drawing { Id = "rId1" }));

            Assert.Equal(
                $@"{XmlDeclaration}<worksheet xmlns=""{SpreadsheetNs}""><sheetData /><drawing r:id=""rId1"" xmlns:r=""{RelationshipsNs}"" /></worksheet>",
                SaveAndRead(worksheetPart));
        }

        [Fact]
        public void DictionaryEntryIsIgnoredForOtherRootNamespaces()
        {
            var settings = new XmlNamespacePrefixSettings();
            settings.Prefixes[WordNs] = string.Empty;

            using var stream = new MemoryStream();
            using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook).UseNamespacePrefixes(settings);
            var worksheetPart = AddWorksheet(doc, new Worksheet(new SheetData()));

            Assert.Equal($@"{XmlDeclaration}<x:worksheet xmlns:x=""{SpreadsheetNs}""><x:sheetData /></x:worksheet>", SaveAndRead(worksheetPart));
        }

        [Fact]
        public void NonEmptyPrefixIsRejected()
        {
            var settings = new XmlNamespacePrefixSettings();
            settings.Prefixes[SpreadsheetNs] = "ss";

            using var stream = new MemoryStream();
            using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);

            Assert.Throws<ArgumentException>(() => doc.UseNamespacePrefixes(settings));
        }

        [Fact]
        public void NonEmptyPrefixIsRejectedOnOpen()
        {
            var settings = new XmlNamespacePrefixSettings();
            settings.Prefixes[SpreadsheetNs] = "ss";

            using var stream = GetStream(TestFiles.Spreadsheet, true);

            Assert.Throws<ArgumentException>(() => SpreadsheetDocument.Open(stream, true, new OpenSettings { NamespacePrefixes = settings }));
        }

        [Fact]
        public void NonEmptyPrefixIsRejectedBeforeFileIsOpened()
        {
            var settings = new XmlNamespacePrefixSettings();
            settings.Prefixes[SpreadsheetNs] = "ss";
            var path = GetTestFilePath(TestFiles.Spreadsheet);

            try
            {
                Assert.Throws<ArgumentException>(() => SpreadsheetDocument.Open(path, true, new OpenSettings { NamespacePrefixes = settings }));

                // the file must not be left open
                using (File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                }
            }
            finally
            {
                File.Delete(path);
            }
        }

        [Fact]
        public void PrefixDeclarationForAttributesIsSkippedWhenPrefixIsUsedByDescendant()
        {
            using var stream = new MemoryStream();
            using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document).UseDefaultNamespaceForRoot();
            var mainPart = doc.AddMainDocumentPart();
            var body = new W.Body(new W.Paragraph(new W.ParagraphProperties(new W.Justification { Val = W.JustificationValues.Center })));
            body.AddNamespaceDeclaration("w", "urn:other");
            mainPart.Document = new W.Document(body);
            mainPart.Document.AddNamespaceDeclaration("r", RelationshipsNs);

            var xml = SaveAndRead(mainPart);

            var jc = System.Xml.Linq.XDocument.Parse(xml).Descendants(System.Xml.Linq.XName.Get("jc", WordNs)).Single();
            Assert.Equal("center", jc.Attribute(System.Xml.Linq.XName.Get("val", WordNs))!.Value);
        }

        [Fact]
        public void CustomFeatureReturningNullPrefixIsTreatedAsNotConfigured()
        {
            using var stream = new MemoryStream();
            using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            doc.Features.Set<DocumentFormat.OpenXml.Features.IXmlNamespacePrefixFeature>(new NullPrefixFeature());
            var worksheetPart = AddWorksheet(doc, new Worksheet(new SheetData()));

            Assert.Equal($@"{XmlDeclaration}<x:worksheet xmlns:x=""{SpreadsheetNs}""><x:sheetData /></x:worksheet>", SaveAndRead(worksheetPart));
        }

        private sealed class NullPrefixFeature : DocumentFormat.OpenXml.Features.IXmlNamespacePrefixFeature
        {
            public bool UseDefaultNamespaceForRoot => false;

            public bool PreserveLoadedDefaultNamespace => false;

            public bool TryGetPrefix(string namespaceUri, out string prefix)
            {
                prefix = null!;
                return true;
            }
        }

        [Fact]
        public void OpenSettingsApplyPresetToLoadedParts()
        {
            using var stream = GetStream(TestFiles.Spreadsheet, true);
            using var doc = SpreadsheetDocument.Open(stream, true, new OpenSettings { NamespacePrefixes = new XmlNamespacePrefixSettings { UseDefaultNamespaceForRoot = true } });
            var worksheetPart = doc.WorkbookPart!.WorksheetParts.First();

            var xml = SaveAndRead(worksheetPart);

            Assert.StartsWith("<worksheet ", StripDeclaration(xml));
            Assert.Contains($@"xmlns=""{SpreadsheetNs}""", xml);
            Assert.DoesNotContain("<x:", xml);
        }

        [Fact]
        public void PartSettingsOverridePackageSettings()
        {
            using var stream = new MemoryStream();
            using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook).UseDefaultNamespaceForRoot();
            var worksheetPart = AddWorksheet(doc, new Worksheet(new SheetData()));
            worksheetPart.UseNamespacePrefixes(new XmlNamespacePrefixSettings());

            Assert.Equal($@"{XmlDeclaration}<x:worksheet xmlns:x=""{SpreadsheetNs}""><x:sheetData /></x:worksheet>", SaveAndRead(worksheetPart));
        }

        [Fact]
        public void OuterXmlMatchesSavedOutput()
        {
            using var stream = new MemoryStream();
            using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook).UseDefaultNamespaceForRoot();
            var cell = new Cell(new CellValue("1"));
            var worksheetPart = AddWorksheet(doc, new Worksheet(new SheetData(new Row(cell))));

            Assert.Equal($@"<worksheet xmlns=""{SpreadsheetNs}""><sheetData><row><c><v>1</v></c></row></sheetData></worksheet>", worksheetPart.Worksheet.OuterXml);
            Assert.Equal($@"<c xmlns=""{SpreadsheetNs}""><v>1</v></c>", cell.OuterXml);
        }

        [Fact]
        public void PrefixIsNotAffectedBySettings()
        {
            using var stream = new MemoryStream();
            using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook).UseDefaultNamespaceForRoot();
            var cell = new Cell(new CellValue("1"));
            var worksheetPart = AddWorksheet(doc, new Worksheet(new SheetData(new Row(cell))));

            Assert.Equal("x", worksheetPart.Worksheet.Prefix);
            Assert.Equal("x", cell.Prefix);
        }

        [Fact]
        public void WriteToAnyXmlWriterMatchesSavedOutput()
        {
            using var stream = new MemoryStream();
            using var doc = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document).UseDefaultNamespaceForRoot();
            var mainPart = doc.AddMainDocumentPart();
            mainPart.Document = new W.Document(new W.Body(new W.Paragraph(new W.ParagraphProperties(new W.Justification { Val = W.JustificationValues.Center }))));

            var sb = new System.Text.StringBuilder();
            using (var writer = System.Xml.XmlWriter.Create(sb, new System.Xml.XmlWriterSettings { OmitXmlDeclaration = true }))
            {
                mainPart.Document.WriteTo(writer);
            }

            Assert.Equal($@"<document xmlns:w=""{WordNs}"" xmlns=""{WordNs}""><body><p><pPr><jc w:val=""center"" /></pPr></p></body></document>", sb.ToString());
        }

        [Fact]
        public void ClonedPackageKeepsSettingsChangedAfterOpen()
        {
            using var stream = GetStream(TestFiles.Spreadsheet, true);
            using var doc = SpreadsheetDocument.Open(stream, true, new OpenSettings { NamespacePrefixes = new XmlNamespacePrefixSettings { UseDefaultNamespaceForRoot = true } });
            doc.UseNamespacePrefixes(new XmlNamespacePrefixSettings());

            using var clone = doc.Clone();

            Assert.StartsWith("<x:worksheet ", StripDeclaration(SaveAndRead(clone.WorkbookPart!.WorksheetParts.First())));
        }

        [Fact]
        public void CloneUsesSettingsPassedToClone()
        {
            using var stream = GetStream(TestFiles.Spreadsheet, true);
            using var doc = SpreadsheetDocument.Open(stream, true);

            using var cloneStream = new MemoryStream();
            using var clone = doc.Clone(cloneStream, true, new OpenSettings { NamespacePrefixes = new XmlNamespacePrefixSettings { UseDefaultNamespaceForRoot = true } });

            Assert.StartsWith("<worksheet ", StripDeclaration(SaveAndRead(clone.WorkbookPart!.WorksheetParts.First())));
        }

        [Fact]
        public void CloneWithSettingsWithoutNamespacePrefixesUsesBuiltInPrefixes()
        {
            using var stream = GetStream(TestFiles.Spreadsheet, true);
            using var doc = SpreadsheetDocument.Open(stream, true, new OpenSettings { NamespacePrefixes = new XmlNamespacePrefixSettings { UseDefaultNamespaceForRoot = true } });

            using var cloneStream = new MemoryStream();
            using var clone = doc.Clone(cloneStream, true, new OpenSettings());

            Assert.StartsWith("<x:worksheet ", StripDeclaration(SaveAndRead(clone.WorkbookPart!.WorksheetParts.First())));
        }

        [Fact]
        public void PresetAppliesToRootCreatedFromOuterXml()
        {
            using var stream = new MemoryStream();
            using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook).UseDefaultNamespaceForRoot();
            var worksheetPart = AddWorksheet(doc, new Worksheet($@"<x:worksheet xmlns:x=""{SpreadsheetNs}""><x:sheetData /></x:worksheet>"));

            Assert.Equal($@"{XmlDeclaration}<worksheet xmlns:x=""{SpreadsheetNs}"" xmlns=""{SpreadsheetNs}""><sheetData /></worksheet>", SaveAndRead(worksheetPart));
        }

        [Fact]
        public void ClonedPackageKeepsSettingsAppliedWithExtension()
        {
            using var stream = new MemoryStream();
            using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook).UseDefaultNamespaceForRoot();
            AddWorksheet(doc, new Worksheet(new SheetData())).Worksheet.Save();

            using var clone = doc.Clone();

            Assert.StartsWith("<worksheet ", StripDeclaration(SaveAndRead(clone.WorkbookPart!.WorksheetParts.First())));
        }

        [Fact]
        public void UntouchedPartsAreNotReserialized()
        {
            byte[] originalBytes;
            using (var resource = GetStream(TestFiles.Spreadsheet))
            using (var buffer = new MemoryStream())
            {
                resource.CopyTo(buffer);
                originalBytes = buffer.ToArray();
            }

            using var edited = new MemoryStream();
            edited.Write(originalBytes, 0, originalBytes.Length);
            edited.Position = 0;

            using (SpreadsheetDocument.Open(edited, true, new OpenSettings { NamespacePrefixes = new XmlNamespacePrefixSettings { UseDefaultNamespaceForRoot = true } }))
            {
            }

            using var originalPackage = Package.Open(new MemoryStream(originalBytes), FileMode.Open, FileAccess.Read);
            edited.Position = 0;
            using var editedPackage = Package.Open(edited, FileMode.Open, FileAccess.Read);

            foreach (var part in originalPackage.GetParts().Where(p => !p.Uri.OriginalString.EndsWith(".rels", StringComparison.Ordinal)))
            {
                Assert.Equal(ReadBytes(part), ReadBytes(editedPackage.GetPart(part.Uri)));
            }
        }

        [Fact]
        public void StrictDocumentUsesTransitionalNamespaceAsDefault()
        {
            using var stream = GetStream(TestFiles.Strict01, true);
            using var doc = WordprocessingDocument.Open(stream, true, new OpenSettings { NamespacePrefixes = new XmlNamespacePrefixSettings { UseDefaultNamespaceForRoot = true } });

            var xml = SaveAndRead(doc.MainDocumentPart!);

            Assert.StartsWith("<document ", StripDeclaration(xml));
            Assert.Contains($@"xmlns=""{WordNs}""", xml);
        }

        [Fact]
        public void StrictNamespaceDictionaryKeyMatchesTransitionalRoot()
        {
            var settings = new XmlNamespacePrefixSettings();
            settings.Prefixes[StrictWordNs] = string.Empty;

            using var stream = GetStream(TestFiles.Strict01, true);
            using var doc = WordprocessingDocument.Open(stream, true, new OpenSettings { NamespacePrefixes = settings });

            Assert.StartsWith("<document ", StripDeclaration(SaveAndRead(doc.MainDocumentPart!)));
        }

        [Fact]
        public void ClonedPackageKeepsSettings()
        {
            using var stream = GetStream(TestFiles.Spreadsheet, true);
            using var doc = SpreadsheetDocument.Open(stream, true, new OpenSettings { NamespacePrefixes = new XmlNamespacePrefixSettings { UseDefaultNamespaceForRoot = true } });
            using var clone = doc.Clone();

            Assert.StartsWith("<worksheet ", StripDeclaration(SaveAndRead(clone.WorkbookPart!.WorksheetParts.First())));
        }

        private static WorksheetPart AddWorksheet(SpreadsheetDocument doc, Worksheet worksheet)
        {
            var workbookPart = doc.AddWorkbookPart();
            workbookPart.Workbook = new Workbook(new Sheets());
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = worksheet;
            return worksheetPart;
        }

        private static string SaveAndRead(OpenXmlPart part)
        {
            part.RootElement!.Save();

            using var reader = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read));
            return reader.ReadToEnd();
        }

        private static string StripDeclaration(string xml)
            => xml.StartsWith("<?xml", StringComparison.Ordinal) ? xml.Substring(xml.IndexOf("?>", StringComparison.Ordinal) + 2) : xml;

        private static byte[] ReadBytes(PackagePart part)
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var ms = new MemoryStream();
            stream.CopyTo(ms);
            return ms.ToArray();
        }
    }
}
