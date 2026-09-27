// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Framework;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;

namespace DocumentFormat.OpenXml.Features
{
    internal sealed class XmlNamespacePrefixFeature : IXmlNamespacePrefixFeature
    {
        private readonly Dictionary<string, string> _prefixes = new();

        public XmlNamespacePrefixFeature(XmlNamespacePrefixSettings settings)
        {
            if (settings is null)
            {
                throw new ArgumentNullException(nameof(settings));
            }

            UseDefaultNamespaceForRoot = settings.UseDefaultNamespaceForRoot;
            PreserveLoadedDefaultNamespace = settings.PreserveLoadedDefaultNamespace;

            var resolver = FeatureCollection.Default.GetNamespaceResolver();

            foreach (var item in settings.Prefixes)
            {
                // Remapping a prefixed namespace would invalidate mc:Ignorable and mc:ProcessContent, which refer to namespaces by prefix
                if (!string.IsNullOrEmpty(item.Value))
                {
                    throw new ArgumentException(SR.Format(ExceptionMessages.NamespacePrefixNotSupported, item.Value, item.Key), nameof(settings));
                }

                // Part content is always written with transitional namespaces, so match entries given for strict namespaces as well
                var uri = resolver.TryGetTransitionalNamespace(new OpenXmlNamespace(item.Key), out var transitional) ? transitional.Uri : item.Key;

                _prefixes[uri] = string.Empty;
            }
        }

        public bool UseDefaultNamespaceForRoot { get; }

        public bool PreserveLoadedDefaultNamespace { get; }

        public bool TryGetPrefix(string namespaceUri, out string prefix)
        {
            if (_prefixes.TryGetValue(namespaceUri, out var result))
            {
                prefix = result;
                return true;
            }

            prefix = string.Empty;
            return false;
        }
    }
}
