// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features;
using System;

namespace DocumentFormat.OpenXml.Packaging
{
    /// <summary>
    /// Extensions to control which namespace prefixes are used when part XML is written.
    /// </summary>
    public static class XmlNamespacePrefixExtensions
    {
        /// <summary>
        /// Uses the given namespace prefix settings when writing the XML of a package's parts, or of a single part.
        /// </summary>
        /// <typeparam name="TContainer">The type of the package or part.</typeparam>
        /// <param name="container">The package or part.</param>
        /// <param name="settings">The settings to use. They are copied, so later changes to them have no effect.</param>
        /// <returns>The <paramref name="container"/>.</returns>
        /// <exception cref="ArgumentException">Thrown if <paramref name="settings"/> contains a prefix that is not supported.</exception>
        public static TContainer UseNamespacePrefixes<TContainer>(this TContainer container, XmlNamespacePrefixSettings settings)
            where TContainer : OpenXmlPartContainer
        {
            if (container is null)
            {
                throw new ArgumentNullException(nameof(container));
            }

            container.Features.Set<IXmlNamespacePrefixFeature>(new XmlNamespacePrefixFeature(settings));

            return container;
        }

        /// <summary>
        /// Writes the namespace of each part's root element as the default namespace (<c>xmlns="..."</c>) instead of with a prefix,
        /// as Office applications do.
        /// </summary>
        /// <typeparam name="TContainer">The type of the package or part.</typeparam>
        /// <param name="container">The package or part.</param>
        /// <returns>The <paramref name="container"/>.</returns>
        public static TContainer UseDefaultNamespaceForRoot<TContainer>(this TContainer container)
            where TContainer : OpenXmlPartContainer
            => container.UseNamespacePrefixes(new XmlNamespacePrefixSettings { UseDefaultNamespaceForRoot = true });
    }
}
