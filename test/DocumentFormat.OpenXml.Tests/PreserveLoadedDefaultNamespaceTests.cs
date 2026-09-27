// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.IO;
using System.Text;
using Xunit;

namespace DocumentFormat.OpenXml.Tests
{
    public class PreserveLoadedDefaultNamespaceTests
    {
        private const string SpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string RelationshipsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        private const string ExcelStyleWorksheet =
            $@"<worksheet xmlns=""{SpreadsheetNs}"" xmlns:r=""{RelationshipsNs}""><sheetData><row r=""1""><c r=""A1""><v>1</v></c></row></sheetData></worksheet>";

        [Fact]
        public void LoadedDefaultNamespaceIsDroppedByDefault()
        {
            var xml = RoundTrip(ExcelStyleWorksheet, settings: null);

            Assert.Equal($@"<x:worksheet xmlns:r=""{RelationshipsNs}"" xmlns:x=""{SpreadsheetNs}""><x:sheetData><x:row r=""1""><x:c r=""A1""><x:v>1</x:v></x:c></x:row></x:sheetData></x:worksheet>", xml);
        }

        [Fact]
        public void LoadedDefaultNamespaceIsPreservedWhenEnabled()
        {
            var xml = RoundTrip(ExcelStyleWorksheet, new XmlNamespacePrefixSettings { PreserveLoadedDefaultNamespace = true });

            Assert.Equal(ExcelStyleWorksheet, xml);
        }

        [Fact]
        public void LoadedPrefixedPartIsUnchangedWhenPreserving()
        {
            const string Prefixed = $@"<x:worksheet xmlns:x=""{SpreadsheetNs}""><x:sheetData /></x:worksheet>";

            var xml = RoundTrip(Prefixed, new XmlNamespacePrefixSettings { PreserveLoadedDefaultNamespace = true });

            Assert.Equal(Prefixed, xml);
        }

        [Fact]
        public void PreservedDefaultForOtherNamespaceKeepsRootPrefixed()
        {
            const string Input = $@"<x:worksheet xmlns:x=""{SpreadsheetNs}"" xmlns=""urn:other""><x:sheetData /></x:worksheet>";

            var xml = RoundTrip(Input, new XmlNamespacePrefixSettings { PreserveLoadedDefaultNamespace = true });

            Assert.Equal(Input, xml);
        }

        [Fact]
        public void ConfiguredDefaultNamespaceWinsOverPreservedDeclaration()
        {
            const string Input = $@"<x:worksheet xmlns:x=""{SpreadsheetNs}"" xmlns=""urn:other""><x:sheetData /></x:worksheet>";

            var xml = RoundTrip(Input, new XmlNamespacePrefixSettings { PreserveLoadedDefaultNamespace = true, UseDefaultNamespaceForRoot = true });

            Assert.Equal($@"<worksheet xmlns:x=""{SpreadsheetNs}"" xmlns=""{SpreadsheetNs}""><sheetData /></worksheet>", xml);
        }

        private static string RoundTrip(string worksheetXml, XmlNamespacePrefixSettings settings)
        {
            using var stream = new MemoryStream();

            using (var doc = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = doc.AddWorkbookPart();
                workbookPart.Workbook = new Workbook(new Sheets());
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();

                using var data = new MemoryStream(Encoding.UTF8.GetBytes(worksheetXml));
                worksheetPart.FeedData(data);
            }

            stream.Position = 0;

            using (var doc = SpreadsheetDocument.Open(stream, true, new OpenSettings { NamespacePrefixes = settings }))
            {
                var worksheetPart = doc.WorkbookPart!.GetPartsOfType<WorksheetPart>().GetEnumerator();
                Assert.True(worksheetPart.MoveNext());

                var part = worksheetPart.Current;
                part.Worksheet.Save();

                using var reader = new StreamReader(part.GetStream(FileMode.Open, FileAccess.Read));
                var xml = reader.ReadToEnd();

                return xml.StartsWith("<?xml", StringComparison.Ordinal) ? xml.Substring(xml.IndexOf("?>", StringComparison.Ordinal) + 2) : xml;
            }
        }
    }
}
