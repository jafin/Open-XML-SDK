// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace DocumentFormat.OpenXml.Features
{
    internal static class XmlNamespacePrefixFeatureExtensions
    {
        /// <summary>
        /// Gets a value indicating whether a part root element in <paramref name="namespaceUri"/> is written in the default namespace.
        /// </summary>
        public static bool IsDefaultNamespaceForRoot(this IXmlNamespacePrefixFeature? feature, string namespaceUri)
        {
            if (feature is null)
            {
                return false;
            }

            // only an empty prefix is currently honoured; a missing prefix falls back to the preset
            if (feature.TryGetPrefix(namespaceUri, out var prefix) && prefix is not null)
            {
                return prefix.Length == 0;
            }

            return feature.UseDefaultNamespaceForRoot;
        }
    }
}
