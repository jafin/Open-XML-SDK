// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
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

        [Fact]
        public void InvalidSettingsLeavePartUntouched()
        {
            using var stream = new MemoryStream();
            using var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            var worksheetPart = AddWorksheetPart(doc);
            worksheetPart.Worksheet = new Worksheet(new SheetData());
            worksheetPart.Worksheet.Save();
            var before = Read(worksheetPart);

            var settings = new XmlNamespacePrefixSettings();
            settings.Prefixes[SpreadsheetNs] = "ss";

            Assert.Throws<ArgumentException>(() => new OpenXmlPartWriter(worksheetPart, new OpenXmlPartWriterSettings { NamespacePrefixes = settings }));
            Assert.Equal(before, Read(worksheetPart));
        }

        [Fact]
        public void WriteElementOfRootUsesWriterSettings()
        {
            using var stream = new MemoryStream();

            using (var writer = new OpenXmlPartWriter(stream, new OpenXmlPartWriterSettings { NamespacePrefixes = new XmlNamespacePrefixSettings { UseDefaultNamespaceForRoot = true } }))
            {
                writer.WriteElement(new Worksheet(new SheetData(new Row())));
            }

            Assert.Equal($@"<worksheet xmlns=""{SpreadsheetNs}""><sheetData><row /></sheetData></worksheet>", Read(stream));
        }

        [Fact]
        public void WriteElementOfConfiguredRootIntoWriterWithoutSettingsIsPrefixed()
        {
            using var package = new MemoryStream();
            using var doc = SpreadsheetDocument.Create(package, SpreadsheetDocumentType.Workbook).UseDefaultNamespaceForRoot();
            var worksheetPart = AddWorksheetPart(doc);
            worksheetPart.Worksheet = new Worksheet(new SheetData(new Row()));

            using var stream = new MemoryStream();

            using (var writer = new OpenXmlPartWriter(stream))
            {
                writer.WriteElement(worksheetPart.Worksheet);
            }

            Assert.Equal($@"<x:worksheet xmlns:x=""{SpreadsheetNs}""><x:sheetData><x:row /></x:sheetData></x:worksheet>", Read(stream));
        }

        [Fact]
        public void ConflictingDefaultDeclarationAttributeOnConfiguredRootIsSkipped()
        {
            using var stream = new MemoryStream();

            using (var writer = new OpenXmlPartWriter(stream, new OpenXmlPartWriterSettings { NamespacePrefixes = new XmlNamespacePrefixSettings { UseDefaultNamespaceForRoot = true } }))
            {
                writer.WriteStartElement(new Worksheet(), new[] { new OpenXmlAttribute(string.Empty, "xmlns", "http://www.w3.org/2000/xmlns/", "urn:other") });
                writer.WriteEndElement();
            }

            Assert.Equal($@"<worksheet xmlns:x=""{SpreadsheetNs}"" xmlns=""{SpreadsheetNs}"" />", Read(stream));
        }

        [Fact]
        public void PrefixDeclarationAttributeOnConfiguredRootIsNotDuplicated()
        {
            using var stream = new MemoryStream();

            using (var writer = new OpenXmlPartWriter(stream, new OpenXmlPartWriterSettings { NamespacePrefixes = new XmlNamespacePrefixSettings { UseDefaultNamespaceForRoot = true } }))
            {
                writer.WriteStartElement(new Worksheet(), new[] { new OpenXmlAttribute("xmlns", "x", "http://www.w3.org/2000/xmlns/", SpreadsheetNs) });
                writer.WriteEndElement();
            }

            Assert.Equal($@"<worksheet xmlns:x=""{SpreadsheetNs}"" xmlns=""{SpreadsheetNs}"" />", Read(stream));
        }

        [Fact]
        public void AttributeWithNullPrefixIsWritten()
        {
            using var stream = new MemoryStream();

            using (var writer = new OpenXmlPartWriter(stream))
            {
                writer.WriteStartElement(new Worksheet(), new[] { new OpenXmlAttribute(null!, "foo", string.Empty, "1") });
                writer.WriteEndElement();
            }

            Assert.Equal($@"<x:worksheet foo=""1"" xmlns:x=""{SpreadsheetNs}"" />", Read(stream));
        }

        [Fact]
        public void AttributesFromSingleUseEnumerableAreWritten()
        {
            using var stream = new MemoryStream();

            using (var writer = new OpenXmlPartWriter(stream))
            {
                writer.WriteStartElement(new Row(), new SingleUse<OpenXmlAttribute>(new OpenXmlAttribute(string.Empty, "r", string.Empty, "1")));
                writer.WriteEndElement();
            }

            Assert.Equal($@"<x:row r=""1"" xmlns:x=""{SpreadsheetNs}"" />", Read(stream));
        }

        [Fact]
        public void WriteElementWithoutSettingsMatchesWriteTo()
        {
            var document = new W.Document(new W.Body(new W.Paragraph(new W.Run(new W.Text(string.Empty) { Space = SpaceProcessingModeValues.Preserve }))));

            var expected = new StringBuilder();
            using (var xmlWriter = System.Xml.XmlWriter.Create(expected, new System.Xml.XmlWriterSettings { OmitXmlDeclaration = true }))
            {
                document.WriteTo(xmlWriter);
            }

            using var stream = new MemoryStream();

            using (var writer = new OpenXmlPartWriter(stream))
            {
                writer.WriteElement(document);
            }

            Assert.Equal(expected.ToString(), Read(stream));
        }

        [Fact]
        public void CopiedElementUnderPreservedDefaultIsPrefixedWhenDefaultIsNotInScope()
        {
            var sourceRow = new Row { RowIndex = 1 };
            var source = new Worksheet(new SheetData(sourceRow));
            source.AddNamespaceDeclaration(string.Empty, SpreadsheetNs);

            using var stream = new MemoryStream();

            using (var writer = new OpenXmlPartWriter(stream))
            {
                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());
                writer.WriteStartElement(sourceRow);
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();
            }

            Assert.Equal($@"<x:worksheet xmlns:x=""{SpreadsheetNs}""><x:sheetData><x:row r=""1"" /></x:sheetData></x:worksheet>", Read(stream));
        }

        private sealed class SingleUse<T> : IEnumerable<T>
        {
            private T[] _items;

            public SingleUse(params T[] items) => _items = items;

            public IEnumerator<T> GetEnumerator()
            {
                var items = _items;
                _items = Array.Empty<T>();
                return ((IEnumerable<T>)items).GetEnumerator();
            }

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
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
