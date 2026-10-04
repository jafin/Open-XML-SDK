// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;

namespace DocumentFormat.OpenXml.Packaging
{
    /// <summary>
    /// Settings that control which namespace prefixes are used when part XML is written.
    /// </summary>
    public class XmlNamespacePrefixSettings
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="XmlNamespacePrefixSettings"/> class.
        /// </summary>
        public XmlNamespacePrefixSettings()
        {
        }

        internal XmlNamespacePrefixSettings(XmlNamespacePrefixSettings other)
        {
            UseDefaultNamespaceForRoot = other.UseDefaultNamespaceForRoot;
            PreserveLoadedDefaultNamespace = other.PreserveLoadedDefaultNamespace;

            foreach (var item in other.Prefixes)
            {
                Prefixes.Add(item.Key, item.Value);
            }
        }

        /// <summary>
        /// Gets or sets a value indicating whether the namespace of each part's root element is written as the default namespace
        /// (<c>xmlns="..."</c>) instead of with a prefix, as Office applications do. The default value is <c>false</c>.
        /// </summary>
        public bool UseDefaultNamespaceForRoot { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether a default namespace declaration found when a part is loaded is preserved when the part is saved.
        /// The default value is <c>false</c>.
        /// </summary>
        public bool PreserveLoadedDefaultNamespace { get; set; }

        /// <summary>
        /// Gets the prefixes to use, keyed by namespace uri. An empty prefix writes the namespace as the default namespace.
        /// </summary>
        /// <remarks>
        /// An entry takes precedence over <see cref="UseDefaultNamespaceForRoot"/>. Only the namespace of a part's root element may be
        /// written as the default namespace; entries for other namespaces are ignored for that part. Only empty prefixes are currently
        /// supported, and other values cause an <see cref="System.ArgumentException"/> when the settings are applied.
        /// </remarks>
        public IDictionary<string, string> Prefixes { get; } = new Dictionary<string, string>();
    }
}
