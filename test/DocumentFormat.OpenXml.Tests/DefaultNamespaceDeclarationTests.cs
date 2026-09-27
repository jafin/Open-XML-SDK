// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Text;
using Xunit;

namespace DocumentFormat.OpenXml.Tests
{
    public class DefaultNamespaceDeclarationTests
    {
        private const string SpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private const string WordNs = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        [Fact]
        public void RootWithDefaultDeclarationIsWrittenWithoutPrefix()
        {
            var worksheet = new Worksheet(new SheetData(new Row(new Cell(new CellValue("1")))));
            worksheet.AddNamespaceDeclaration(string.Empty, SpreadsheetNs);

            Assert.Equal($@"<worksheet xmlns=""{SpreadsheetNs}""><sheetData><row><c><v>1</v></c></row></sheetData></worksheet>", worksheet.OuterXml);
        }

        [Fact]
        public void SaveWritesDefaultDeclarationWithoutPrefix()
        {
            var worksheet = new Worksheet(new SheetData());
            worksheet.AddNamespaceDeclaration(string.Empty, SpreadsheetNs);

            using var stream = new MemoryStream();
            worksheet.Save(stream);

            Assert.Equal($@"<?xml version=""1.0"" encoding=""utf-8""?><worksheet xmlns=""{SpreadsheetNs}""><sheetData /></worksheet>", Encoding.UTF8.GetString(stream.ToArray()));
        }

        [Fact]
        public void ElementsUseDefaultNamespaceWhenPrefixedDeclarationForSameUriIsAlsoInScope()
        {
            var document = new Document(new Body(new Paragraph(new ParagraphProperties(new Justification { Val = JustificationValues.Center }))));
            document.AddNamespaceDeclaration(string.Empty, WordNs);
            document.AddNamespaceDeclaration("w", WordNs);

            Assert.Equal($@"<document xmlns=""{WordNs}"" xmlns:w=""{WordNs}""><body><p><pPr><jc w:val=""center"" /></pPr></p></body></document>", document.OuterXml);
        }

        [Fact]
        public void ElementDeclaringDifferentDefaultNamespaceKeepsPrefixForItself()
        {
            var body = new Body(new Paragraph());
            body.AddNamespaceDeclaration(string.Empty, "urn:other");
            var document = new Document(body);
            document.AddNamespaceDeclaration(string.Empty, WordNs);
            document.AddNamespaceDeclaration("w", WordNs);

            Assert.Equal($@"<document xmlns=""{WordNs}"" xmlns:w=""{WordNs}""><w:body xmlns=""urn:other""><w:p /></w:body></document>", document.OuterXml);
        }

        [Fact]
        public void PrefixIsEmptyUnderDefaultDeclaration()
        {
            var cell = new Cell();
            var worksheet = new Worksheet(new SheetData(new Row(cell)));
            worksheet.AddNamespaceDeclaration(string.Empty, SpreadsheetNs);

            Assert.Equal(string.Empty, worksheet.Prefix);
            Assert.Equal(string.Empty, cell.Prefix);
            Assert.Equal(SpreadsheetNs, cell.LookupNamespace(string.Empty));
        }

        [Fact]
        public void RemoveNamespaceDeclarationRemovesDefaultDeclaration()
        {
            var worksheet = new Worksheet();
            worksheet.AddNamespaceDeclaration(string.Empty, SpreadsheetNs);

            worksheet.RemoveNamespaceDeclaration(string.Empty);

            Assert.Empty(worksheet.NamespaceDeclarations);
            Assert.Equal("x", worksheet.Prefix);
        }

        [Fact]
        public void AddingDefaultDeclarationTwiceThrows()
        {
            var worksheet = new Worksheet();
            worksheet.AddNamespaceDeclaration(string.Empty, SpreadsheetNs);

            Assert.Throws<InvalidOperationException>(() => worksheet.AddNamespaceDeclaration(string.Empty, SpreadsheetNs));
        }

        [Fact]
        public void NullPrefixStillThrows()
        {
            Assert.Throws<ArgumentNullException>(() => new Worksheet().AddNamespaceDeclaration(null!, SpreadsheetNs));
        }

        [Fact]
        public void OpenXmlWriterWritesDefaultDeclaration()
        {
            var worksheet = new Worksheet();
            worksheet.AddNamespaceDeclaration(string.Empty, SpreadsheetNs);

            using var stream = new MemoryStream();
            using (var writer = OpenXmlWriter.Create(stream))
            {
                writer.WriteStartElement(worksheet);
                writer.WriteEndElement();
            }

            Assert.Equal($@"<?xml version=""1.0"" encoding=""utf-8""?><worksheet xmlns=""{SpreadsheetNs}"" />", Encoding.UTF8.GetString(stream.ToArray()).TrimStart('﻿'));
        }
    }
}
