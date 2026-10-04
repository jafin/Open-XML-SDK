// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Text;
using Xunit;

namespace DocumentFormat.OpenXml.Tests
{
    public class OpenXmlPartRootElementPrefixTests
    {
        private const string SpreadsheetNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        [Fact]
        public void RootUsesPrefixDeclaredOnItselfForItsOwnNamespace()
        {
            var worksheet = new Worksheet(new SheetData());
            worksheet.AddNamespaceDeclaration("ss", SpreadsheetNs);

            Assert.Equal($@"<ss:worksheet xmlns:ss=""{SpreadsheetNs}""><ss:sheetData /></ss:worksheet>", worksheet.OuterXml);
        }

        [Fact]
        public void SaveUsesPrefixDeclaredOnRootForItsOwnNamespace()
        {
            var worksheet = new Worksheet(new SheetData());
            worksheet.AddNamespaceDeclaration("ss", SpreadsheetNs);

            using var stream = new MemoryStream();
            worksheet.Save(stream);

            Assert.Equal($@"<?xml version=""1.0"" encoding=""utf-8""?><ss:worksheet xmlns:ss=""{SpreadsheetNs}""><ss:sheetData /></ss:worksheet>", Encoding.UTF8.GetString(stream.ToArray()));
        }
    }
}
