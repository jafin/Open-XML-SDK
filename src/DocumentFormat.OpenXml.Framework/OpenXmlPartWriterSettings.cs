// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Packaging;
using System.Text;

namespace DocumentFormat.OpenXml;

/// <summary>
/// Settings for the <see cref="OpenXmlPartWriter" /> .
/// </summary>
public class OpenXmlPartWriterSettings
{
#if FEATURE_ASYNC_SAX_XML
    /// <summary>
    /// Gets or sets a value indicating whether asynchronous <see cref="OpenXmlPartWriter" />  methods can be used.
    /// </summary>
    public bool Async { get; set; }
#endif

    /// <summary>
    /// Gets or sets a value indicating whether the <see cref="OpenXmlPartWriter" />  should also close the underlying stream or TextWriter when the Close() method is called.
    /// </summary>
    public bool CloseOutput { get; set; }

    /// <summary>
    /// Gets or sets the type of text encoding to use.
    /// </summary>
    public Encoding Encoding { get; set; } = Encoding.UTF8;

    /// <summary>
    /// Gets or sets the settings that control which namespace prefixes are written. When <c>null</c> (the default), a writer
    /// created for a part uses the settings of the part or its package, and a writer created for a stream uses the built-in prefixes.
    /// </summary>
    /// <remarks>
    /// When the root element is written in the default namespace, the built-in prefix of that namespace is declared on the root as
    /// well, because the elements that follow cannot be inspected in advance to tell whether an attribute needs it.
    /// </remarks>
    public XmlNamespacePrefixSettings? NamespacePrefixes { get; set; }
}
