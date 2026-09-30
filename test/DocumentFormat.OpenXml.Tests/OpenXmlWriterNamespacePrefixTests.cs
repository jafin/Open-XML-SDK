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

        [Fact]
        public void CopyingWithReaderWithoutSettingsKeepsPrefixes()
        {
            var xml = CopyExcelStyleSheet(settings: null);

            Assert.StartsWith("<x:worksheet ", xml);
            Assert.Contains("<x:sheetData><x:row r=\"1\"><x:c r=\"A1\"><x:v>1</x:v></x:c></x:row></x:sheetData>", xml);
        }

        [Fact]
        public void CopyingWithReaderWithSettingsUsesDefaultNamespace()
        {
            var xml = CopyExcelStyleSheet(new XmlNamespacePrefixSettings { UseDefaultNamespaceForRoot = true });

            Assert.StartsWith("<worksheet ", xml);
            Assert.Contains("<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>", xml);
        }

        [Fact]
        public void StreamedElementRedeclaringDefaultNamespaceGetsGeneratedPrefixWhenNoneIsKnown()
        {
            var element = new OpenXmlUnknownElement(string.Empty, "item", "urn:custom");
            element.AddNamespaceDeclaration(string.Empty, "urn:other");

            using var stream = new MemoryStream();

            using (var writer = new OpenXmlPartWriter(stream))
            {
                writer.WriteStartElement(element, Array.Empty<OpenXmlAttribute>());
                writer.WriteEndElement();
            }

            Assert.Equal(@"<ns0:item xmlns=""urn:other"" xmlns:ns0=""urn:custom"" />", Read(stream));
        }

        [Fact]
        public void UnknownElementWithEmptyPrefixIsUnchangedWithoutSettings()
        {
            using var stream = new MemoryStream();

            using (var writer = new OpenXmlPartWriter(stream))
            {
                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new OpenXmlUnknownElement(string.Empty, "shape", VmlNs));
                writer.WriteEndElement();
                writer.WriteEndElement();
            }

            Assert.Equal($@"<x:worksheet xmlns:x=""{SpreadsheetNs}""><shape xmlns=""{VmlNs}"" /></x:worksheet>", Read(stream));
        }

        [Fact]
        public void CopiedUnknownElementDeclaringItsNamespaceIsUnchangedWithoutSettings()
        {
            var xml = Copy($@"<x:worksheet xmlns:x=""{SpreadsheetNs}""><x:sheetData /><shape xmlns=""{VmlNs}"" /></x:worksheet>", settings: null);

            Assert.Equal($@"<x:worksheet xmlns:x=""{SpreadsheetNs}""><x:sheetData /><shape xmlns=""{VmlNs}"" /></x:worksheet>", xml);
        }

        [Fact]
        public void CopyingRowsAsElementsWithoutSettingsKeepsPrefixes()
        {
            var xml = CopyExcelStyleSheet(settings: null, loadRows: true);

            Assert.Contains("<x:sheetData><x:row r=\"1\"><x:c r=\"A1\"><x:v>1</x:v></x:c></x:row></x:sheetData>", xml);
        }

        [Fact]
        public void CopyingRowsAsElementsWithSettingsUsesDefaultNamespace()
        {
            var xml = CopyExcelStyleSheet(new XmlNamespacePrefixSettings { UseDefaultNamespaceForRoot = true }, loadRows: true);

            Assert.Contains("<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>", xml);
        }

        [Fact]
        public void CopyingWithOnlyPreserveLoadedDefaultNamespaceKeepsPrefixes()
        {
            var settings = new XmlNamespacePrefixSettings { PreserveLoadedDefaultNamespace = true };

            Assert.Contains("<x:sheetData><x:row r=\"1\">", CopyExcelStyleSheet(settings));
            Assert.Contains("<x:sheetData><x:row r=\"1\">", CopyExcelStyleSheet(settings, loadRows: true));
        }

        [Fact]
        public void AttributeRedeclaringDefaultNamespaceUnderDeclaredDefaultKeepsPrefix()
        {
            var worksheet = new Worksheet();
            worksheet.AddNamespaceDeclaration(string.Empty, SpreadsheetNs);

            using var stream = new MemoryStream();

            using (var writer = new OpenXmlPartWriter(stream))
            {
                writer.WriteStartElement(worksheet);
                writer.WriteStartElement(new SheetData(), new[] { new OpenXmlAttribute(string.Empty, "xmlns", "http://www.w3.org/2000/xmlns/", "urn:other") });
                writer.WriteElement(new Row());
                writer.WriteEndElement();
                writer.WriteEndElement();
            }

            Assert.Equal($@"<worksheet xmlns=""{SpreadsheetNs}""><x:sheetData xmlns=""urn:other"" xmlns:x=""{SpreadsheetNs}""><x:row /></x:sheetData></worksheet>", Read(stream));
        }

        [Fact]
        public void UnparsedElementUnderDefaultNamespaceKeepsChildrenOutOfDefaultNamespace()
        {
            using var stream = new MemoryStream();

            using (var writer = new OpenXmlPartWriter(stream, new OpenXmlPartWriterSettings { NamespacePrefixes = new XmlNamespacePrefixSettings { UseDefaultNamespaceForRoot = true } }))
            {
                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());
                writer.WriteElement(new Row($@"<x:row xmlns:x=""{SpreadsheetNs}""><foo /></x:row>"));
                writer.WriteEndElement();
                writer.WriteEndElement();
            }

            Assert.Contains(@"<foo xmlns="""">", Read(stream));
        }

        [Fact]
        public void WriteElementUnderDefaultNamespaceMatchesWriteTo()
        {
            var body = new W.Body(new W.Paragraph(new W.Run(new W.Text(string.Empty) { Space = SpaceProcessingModeValues.Preserve })));

            using var stream = new MemoryStream();

            using (var writer = new OpenXmlPartWriter(stream, new OpenXmlPartWriterSettings { NamespacePrefixes = new XmlNamespacePrefixSettings { UseDefaultNamespaceForRoot = true } }))
            {
                writer.WriteStartElement(new W.Document());
                writer.WriteElement(body);
                writer.WriteEndElement();
            }

            Assert.Contains(@"<t xml:space=""preserve""></t>", Read(stream));
        }

        private const string VmlNs = "urn:schemas-microsoft-com:vml";

        private static string CopyExcelStyleSheet(XmlNamespacePrefixSettings settings, bool loadRows = false)
            => Copy($@"<worksheet xmlns=""{SpreadsheetNs}"" xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""><sheetData><row r=""1""><c r=""A1""><v>1</v></c></row></sheetData></worksheet>", settings, loadRows);

        private static string Copy(string sourceXml, XmlNamespacePrefixSettings settings, bool loadRows = false)
        {
            using var package = new MemoryStream();
            using var doc = SpreadsheetDocument.Create(package, SpreadsheetDocumentType.Workbook);
            var sourcePart = AddWorksheetPart(doc);

            using (var data = new MemoryStream(Encoding.UTF8.GetBytes(sourceXml)))
            {
                sourcePart.FeedData(data);
            }

            using var target = new MemoryStream();

            using (var reader = new OpenXmlPartReader(sourcePart))
            using (var writer = new OpenXmlPartWriter(target, new OpenXmlPartWriterSettings { NamespacePrefixes = settings }))
            {
                while (reader.Read())
                {
                    if (loadRows && reader.IsStartElement && reader.ElementType == typeof(Row))
                    {
                        // the common streaming pattern: copy the start tags, and load and write each row as an element
                        writer.WriteElement(reader.LoadCurrentElement()!);
                    }
                    else if (reader.IsStartElement)
                    {
                        writer.WriteStartElement(reader);

                        if (reader.ElementType.IsSubclassOf(typeof(OpenXmlLeafTextElement)))
                        {
                            writer.WriteString(reader.GetText());
                        }
                    }
                    else if (reader.IsEndElement)
                    {
                        writer.WriteEndElement();
                    }
                }
            }

            return Read(target);
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
