// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Xml;

#if FEATURE_ASYNC_SAX_XML
using System.Threading.Tasks;
#endif

namespace DocumentFormat.OpenXml
{
    internal static class XmlWriterNamespaceExtensions
    {
        /// <summary>
        /// Writes a namespace declaration. An empty <paramref name="prefix"/> writes the default namespace declaration (<c>xmlns="uri"</c>).
        /// </summary>
        public static void WriteNamespaceDeclaration(this XmlWriter xmlWriter, string prefix, string uri)
        {
            if (prefix.Length == 0)
            {
                xmlWriter.WriteAttributeString(OpenXmlElementContext.XmlnsPrefix, OpenXmlElementContext.XmlnsUri, uri);
            }
            else
            {
                xmlWriter.WriteAttributeString(OpenXmlElementContext.XmlnsPrefix, prefix, OpenXmlElementContext.XmlnsUri, uri);
            }
        }

#if FEATURE_ASYNC_SAX_XML
        /// <summary>
        /// Writes a namespace declaration. An empty <paramref name="prefix"/> writes the default namespace declaration (<c>xmlns="uri"</c>).
        /// </summary>
        public static Task WriteNamespaceDeclarationAsync(this XmlWriter xmlWriter, string prefix, string uri)
            => prefix.Length == 0
                ? xmlWriter.WriteAttributeStringAsync(null, OpenXmlElementContext.XmlnsPrefix, OpenXmlElementContext.XmlnsUri, uri)
                : xmlWriter.WriteAttributeStringAsync(OpenXmlElementContext.XmlnsPrefix, prefix, OpenXmlElementContext.XmlnsUri, uri);
#endif
    }
}
