// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Xunit;

using W = DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentFormat.OpenXml.Tests
{
    public class OpenXmlWriterNamespacePrefixTests
    {
        private const string SpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string WordNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        [Fact]
        public void PartWriterUsesPackageSettings()
        {
            using var stream = new MemoryStream();
            using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook).UseDefaultNamespaceForRoot();
            var worksheetPart = AddWorksheetPart(doc);

            using (var writer = OpenXmlWriter.Create(worksheetPart))
            {
                WriteSheet(writer);
            }

            Assert.Equal($@"<worksheet xmlns:x=""{SpreadsheetNs}"" xmlns=""{SpreadsheetNs}""><sheetData><row><c><v>1</v></c></row></sheetData></worksheet>", Read(worksheetPart));
        }

        [Fact]
        public void WriterSettingsOverridePackageSettings()
        {
            using var stream = new MemoryStream();
            using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook).UseDefaultNamespaceForRoot();
            var worksheetPart = AddWorksheetPart(doc);

            using (var writer = new OpenXmlPartWriter(worksheetPart, new OpenXmlPartWriterSettings { NamespacePrefixes = new XmlNamespacePrefixSettings() }))
            {
                WriteSheet(writer);
            }

            Assert.Equal($@"<x:worksheet xmlns:x=""{SpreadsheetNs}""><x:sheetData><x:row><x:c><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>", Read(worksheetPart));
        }

        [Fact]
        public void StreamWriterUsesSettings()
        {
            using var stream = new MemoryStream();

            using (var writer = new OpenXmlPartWriter(stream, new OpenXmlPartWriterSettings { NamespacePrefixes = new XmlNamespacePrefixSettings { UseDefaultNamespaceForRoot = true } }))
            {
                WriteSheet(writer);
            }

            Assert.Equal($@"<worksheet xmlns:x=""{SpreadsheetNs}"" xmlns=""{SpreadsheetNs}""><sheetData><row><c><v>1</v></c></row></sheetData></worksheet>", Read(stream));
        }

        [Fact]
        public void StreamWriterWithoutSettingsIsUnchanged()
        {
            using var stream = new MemoryStream();

            using (var writer = new OpenXmlPartWriter(stream))
            {
                WriteSheet(writer);
            }

            Assert.Equal($@"<x:worksheet xmlns:x=""{SpreadsheetNs}""><x:sheetData><x:row><x:c><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>", Read(stream));
        }

        [Fact]
        public void AttributesInDefaultNamespaceUseRootPrefixDeclaration()
        {
            using var stream = new MemoryStream();

            using (var writer = new OpenXmlPartWriter(stream, new OpenXmlPartWriterSettings { NamespacePrefixes = new XmlNamespacePrefixSettings { UseDefaultNamespaceForRoot = true } }))
            {
                writer.WriteStartElement(new W.Document());
                writer.WriteStartElement(new W.Body());
                writer.WriteElement(new W.Paragraph(new W.ParagraphProperties(new W.Justification { Val = W.JustificationValues.Center })));
                writer.WriteStartElement(new W.Paragraph(), new[] { new OpenXmlAttribute("w", "rsidR", WordNs, "00000001") });
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();
            }

            Assert.Equal($@"<document xmlns:w=""{WordNs}"" xmlns=""{WordNs}""><body><p><pPr><jc w:val=""center"" /></pPr></p><p w:rsidR=""00000001"" /></body></document>", Read(stream));
        }

        [Fact]
        public void ElementDeclaringDefaultNamespaceIsFollowedByUnprefixedChildren()
        {
            var worksheet = new Worksheet();
            worksheet.AddNamespaceDeclaration(string.Empty, SpreadsheetNs);

            using var stream = new MemoryStream();

            using (var writer = new OpenXmlPartWriter(stream))
            {
                writer.WriteStartElement(worksheet);
                writer.WriteStartElement(new SheetData());
                writer.WriteElement(new Row(new Cell(new CellValue("1"))));
                writer.WriteEndElement();
                writer.WriteEndElement();
            }

            Assert.Equal($@"<worksheet xmlns=""{SpreadsheetNs}""><sheetData><row><c><v>1</v></c></row></sheetData></worksheet>", Read(stream));
        }

#if FEATURE_ASYNC_SAX_XML
        [Fact]
        public async Task AsyncWriterUsesSettings()
        {
            using var stream = new MemoryStream();

            using (var writer = new OpenXmlPartWriter(stream, new OpenXmlPartWriterSettings { Async = true, NamespacePrefixes = new XmlNamespacePrefixSettings { UseDefaultNamespaceForRoot = true } }))
            {
                await writer.WriteStartElementAsync(new Worksheet());
                await writer.WriteStartElementAsync(new SheetData());
                await writer.WriteElementAsync(new Row());
                await writer.WriteEndElementAsync();
                await writer.WriteEndElementAsync();
            }

            Assert.Equal($@"<worksheet xmlns:x=""{SpreadsheetNs}"" xmlns=""{SpreadsheetNs}""><sheetData><row /></sheetData></worksheet>", Read(stream));
        }
#endif

        private static void WriteSheet(OpenXmlWriter writer)
        {
            writer.WriteStartElement(new Worksheet());
            writer.WriteStartElement(new SheetData());
            writer.WriteElement(new Row(new Cell(new CellValue("1"))));
            writer.WriteEndElement();
            writer.WriteEndElement();
        }

        private static WorksheetPart AddWorksheetPart(SpreadsheetDocument doc)
        {
            var workbookPart = doc.AddWorkbookPart();
            workbookPart.Workbook = new Workbook(new Sheets());
            return workbookPart.AddNewPart<WorksheetPart>();
        }

        private static string Read(OpenXmlPart part)
        {
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            return Read(stream);
        }

        private static string Read(Stream stream)
        {
            stream.Position = 0;
            using var reader = new StreamReader(stream, Encoding.UTF8, true, 1024, leaveOpen: true);
            var xml = reader.ReadToEnd();
            return xml.StartsWith("<?xml", StringComparison.Ordinal) ? xml.Substring(xml.IndexOf("?>", StringComparison.Ordinal) + 2) : xml;
        }
    }
}
