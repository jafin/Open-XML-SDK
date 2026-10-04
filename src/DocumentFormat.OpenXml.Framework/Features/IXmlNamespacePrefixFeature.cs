// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace DocumentFormat.OpenXml.Features
{
    /// <summary>
    /// A feature that controls which namespace prefixes are used when part XML is written.
    /// </summary>
    /// <remarks>
    /// Register it on a package to apply it to every part, or on a part to override the package setting for that part.
    /// </remarks>
    public interface IXmlNamespacePrefixFeature
    {
        /// <summary>
        /// Gets a value indicating whether the namespace of each part's root element is written as the default namespace
        /// (<c>xmlns="..."</c>) instead of with a prefix, as Office applications do.
        /// </summary>
        bool UseDefaultNamespaceForRoot { get; }

        /// <summary>
        /// Gets a value indicating whether a default namespace declaration found when a part is loaded is preserved when the part is saved.
        /// </summary>
        bool PreserveLoadedDefaultNamespace { get; }

        /// <summary>
        /// Gets the prefix configured for a namespace. An empty prefix means the namespace is written as the default namespace.
        /// </summary>
        /// <param name="namespaceUri">The namespace uri.</param>
        /// <param name="prefix">The configured prefix.</param>
        /// <returns><c>true</c> if a prefix is configured for <paramref name="namespaceUri"/>; otherwise <c>false</c>.</returns>
        /// <remarks>
        /// An entry takes precedence over <see cref="UseDefaultNamespaceForRoot"/>. Only the namespace of a part's root element may be written as the default namespace.
        /// Only an empty prefix is currently honoured: an entry with any other prefix leaves the namespace written with its built-in prefix.
        /// </remarks>
        bool TryGetPrefix(string namespaceUri, out string prefix);
    }
}
