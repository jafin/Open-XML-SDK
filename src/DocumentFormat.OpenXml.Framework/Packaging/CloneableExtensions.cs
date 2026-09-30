// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Builder;
using DocumentFormat.OpenXml.Features;
using System;
using System.IO;
using System.IO.Packaging;

namespace DocumentFormat.OpenXml.Packaging;

/// <summary>
/// Extensions to enable package cloning.
/// </summary>
public static class CloneableExtensions
{
    /// <summary>
    /// Creates an editable clone of this OpenXml package, opened on a
    /// <see cref="MemoryStream"/> with expandable capacity and using
    /// default OpenSettings. The namespace prefix settings in effect for this package are kept.
    /// </summary>
    /// <returns>The cloned OpenXml package.</returns>
    public static TPackage Clone<TPackage>(this TPackage openXmlPackage)
        where TPackage : OpenXmlPackage
    {
        if (openXmlPackage is null)
        {
            throw new ArgumentNullException(nameof(openXmlPackage));
        }

        return openXmlPackage.CloneOnStream(new MemoryStream(), true, new OpenSettings(), inheritNamespacePrefixes: true);
    }

    /// <summary>
    /// Creates a clone of this OpenXml package, opened on the given stream.
    /// The cloned OpenXml package is opened with the same settings, i.e.,
    /// FileOpenAccess and OpenSettings, as this OpenXml package.
    /// </summary>
    /// <param name="openXmlPackage"></param>
    /// <param name="stream">The IO stream on which to open the OpenXml package.</param>
    /// <returns>The cloned OpenXml package.</returns>
    public static TPackage Clone<TPackage>(this TPackage openXmlPackage, Stream stream)
        where TPackage : OpenXmlPackage
    {
        if (openXmlPackage is null)
        {
            throw new ArgumentNullException(nameof(openXmlPackage));
        }

        if (stream is null)
        {
            throw new ArgumentNullException(nameof(stream));
        }

        return openXmlPackage.CloneOnStream(stream, openXmlPackage.FileOpenAccess == FileAccess.ReadWrite, openXmlPackage.OpenSettings, inheritNamespacePrefixes: true);
    }

    /// <summary>
    /// Creates a clone of this OpenXml package, opened on the given stream.
    /// The cloned OpenXml package is opened with the same OpenSettings as
    /// this OpenXml package.
    /// </summary>
    /// <param name="openXmlPackage"></param>
    /// <param name="stream">The IO stream on which to open the OpenXml package.</param>
    /// <param name="isEditable">In ReadWrite mode. False for Read only mode.</param>
    /// <returns>The cloned OpenXml package.</returns>
    public static TPackage Clone<TPackage>(this TPackage openXmlPackage, Stream stream, bool isEditable)
        where TPackage : OpenXmlPackage
    {
        if (openXmlPackage is null)
        {
            throw new ArgumentNullException(nameof(openXmlPackage));
        }

        if (stream is null)
        {
            throw new ArgumentNullException(nameof(stream));
        }

        return openXmlPackage.CloneOnStream(stream, isEditable, openXmlPackage.OpenSettings, inheritNamespacePrefixes: true);
    }

    /// <summary>
    /// Creates a clone of this OpenXml package, opened on the given stream.
    /// </summary>
    /// <param name="openXmlPackage"></param>
    /// <param name="stream">The IO stream on which to open the OpenXml package.</param>
    /// <param name="isEditable">In ReadWrite mode. False for Read only mode.</param>
    /// <param name="openSettings">The advanced settings for opening a document.</param>
    /// <returns>The cloned OpenXml package.</returns>
    public static TPackage Clone<TPackage>(this TPackage openXmlPackage, Stream stream, bool isEditable, OpenSettings openSettings)
        where TPackage : OpenXmlPackage
    {
        if (openXmlPackage is null)
        {
            throw new ArgumentNullException(nameof(openXmlPackage));
        }

        if (stream is null)
        {
            throw new ArgumentNullException(nameof(stream));
        }

        return openXmlPackage.CloneOnStream(stream, isEditable, openSettings, inheritNamespacePrefixes: false);
    }

    private static TPackage CloneOnStream<TPackage>(this TPackage openXmlPackage, Stream stream, bool isEditable, OpenSettings openSettings, bool inheritNamespacePrefixes)
        where TPackage : OpenXmlPackage
        => openXmlPackage.Features.GetRequired<IPackageFactoryFeature<TPackage>>()
            .Create()
            .UseSettings(ForClone(openSettings, inheritNamespacePrefixes))
            .Build()
            .Open(stream, PackageOpenMode.Create)
            .CopyFrom(openXmlPackage, inheritNamespacePrefixes)
            .Reload(isEditable);

    internal static void Clone<TPackage>(this TPackage source, TPackage destination)
        where TPackage : OpenXmlPackage
        => destination.CopyFrom(source, inheritNamespacePrefixes: true).Reload();

    /// <summary>
    /// Creates a clone of this OpenXml package opened from the given file
    /// (which will be created by cloning this OpenXml package).
    /// The cloned OpenXml package is opened with the same settings, i.e.,
    /// FileOpenAccess and OpenSettings, as this OpenXml package.
    /// </summary>
    /// <param name="openXmlPackage"></param>
    /// <param name="path">The path and file name of the target document.</param>
    /// <returns>The cloned document.</returns>
    public static TPackage Clone<TPackage>(this TPackage openXmlPackage, string path)
        where TPackage : OpenXmlPackage
    {
        if (openXmlPackage is null)
        {
            throw new ArgumentNullException(nameof(openXmlPackage));
        }

        if (string.IsNullOrEmpty(path))
        {
            throw new ArgumentException($"'{nameof(path)}' cannot be null or empty.", nameof(path));
        }

        return openXmlPackage.CloneOnPath(path, openXmlPackage.FileOpenAccess == FileAccess.ReadWrite, openXmlPackage.OpenSettings, inheritNamespacePrefixes: true);
    }

    /// <summary>
    /// Creates a clone of this OpenXml package opened from the given file
    /// (which will be created by cloning this OpenXml package).
    /// The cloned OpenXml package is opened with the same OpenSettings as
    /// this OpenXml package.
    /// </summary>
    /// <param name="openXmlPackage"></param>
    /// <param name="path">The path and file name of the target document.</param>
    /// <param name="isEditable">In ReadWrite mode. False for Read only mode.</param>
    /// <returns>The cloned document.</returns>
    public static TPackage Clone<TPackage>(this TPackage openXmlPackage, string path, bool isEditable)
        where TPackage : OpenXmlPackage
    {
        if (openXmlPackage is null)
        {
            throw new ArgumentNullException(nameof(openXmlPackage));
        }

        if (string.IsNullOrEmpty(path))
        {
            throw new ArgumentException($"'{nameof(path)}' cannot be null or empty.", nameof(path));
        }

        return openXmlPackage.CloneOnPath(path, isEditable, openXmlPackage.OpenSettings, inheritNamespacePrefixes: true);
    }

    /// <summary>
    /// Creates a clone of this OpenXml package opened from the given file (which
    /// will be created by cloning this OpenXml package).
    /// </summary>
    /// <param name="openXmlPackage"></param>
    /// <param name="path">The path and file name of the target document.</param>
    /// <param name="isEditable">In ReadWrite mode. False for Read only mode.</param>
    /// <param name="openSettings">The advanced settings for opening a document.</param>
    /// <returns>The cloned document.</returns>
    public static TPackage Clone<TPackage>(this TPackage openXmlPackage, string path, bool isEditable, OpenSettings? openSettings)
        where TPackage : OpenXmlPackage
    {
        if (openXmlPackage is null)
        {
            throw new ArgumentNullException(nameof(openXmlPackage));
        }

        if (path is null)
        {
            throw new ArgumentNullException(nameof(path));
        }

        return openXmlPackage.CloneOnPath(path, isEditable, openSettings, inheritNamespacePrefixes: false);
    }

    private static TPackage CloneOnPath<TPackage>(this TPackage openXmlPackage, string path, bool isEditable, OpenSettings? openSettings, bool inheritNamespacePrefixes)
        where TPackage : OpenXmlPackage
        => openXmlPackage.Features.GetRequired<IPackageFactoryFeature<TPackage>>()
              .Create()
              .UseSettings(ForClone(openSettings ?? new(), inheritNamespacePrefixes))
              .Build()
              .Open(path, PackageOpenMode.Create)
              .CopyFrom(openXmlPackage, inheritNamespacePrefixes)
              .Reload(isEditable);

    /// <summary>
    /// Creates a clone of this OpenXml package, opened on the specified instance
    /// of Package. The clone will be opened with the same OpenSettings as this
    /// OpenXml package.
    /// </summary>
    /// <param name="openXmlPackage"></param>
    /// <param name="package">The specified instance of Package.</param>
    /// <returns>The cloned OpenXml package.</returns>
    public static TPackage Clone<TPackage>(this TPackage openXmlPackage, Package package)
        where TPackage : OpenXmlPackage
    {
        if (openXmlPackage is null)
        {
            throw new ArgumentNullException(nameof(openXmlPackage));
        }

        if (package is null)
        {
            throw new ArgumentNullException(nameof(package));
        }

        return openXmlPackage.CloneOnPackage(package, openXmlPackage.OpenSettings, inheritNamespacePrefixes: true);
    }

    /// <summary>
    /// Creates a clone of this OpenXml package, opened on the specified instance
    /// of Package.
    /// </summary>
    /// <param name="openXmlPackage"></param>
    /// <param name="package">The specified instance of Package.</param>
    /// <param name="openSettings">The advanced settings for opening a document.</param>
    /// <returns>The cloned OpenXml package.</returns>
    public static TPackage Clone<TPackage>(this TPackage openXmlPackage, Package package, OpenSettings openSettings)
        where TPackage : OpenXmlPackage
    {
        if (openXmlPackage is null)
        {
            throw new ArgumentNullException(nameof(openXmlPackage));
        }

        if (package is null)
        {
            throw new ArgumentNullException(nameof(package));
        }

        return openXmlPackage.CloneOnPackage(package, openSettings, inheritNamespacePrefixes: false);
    }

    private static TPackage CloneOnPackage<TPackage>(this TPackage openXmlPackage, Package package, OpenSettings openSettings, bool inheritNamespacePrefixes)
        where TPackage : OpenXmlPackage
        => openXmlPackage.Features.GetRequired<IPackageFactoryFeature<TPackage>>()
              .Create()
              .UseSettings(ForClone(openSettings ?? new(), inheritNamespacePrefixes))
              .Build()
              .Open(package)
              .CopyFrom(openXmlPackage, inheritNamespacePrefixes);

    private static TPackage CopyFrom<TPackage>(this TPackage destination, TPackage source, bool inheritNamespacePrefixes)
        where TPackage : OpenXmlPackage
    {
        lock (source.Features.GetRequired<ILockFeature>().SyncLock)
        {
            var existing = destination.Features.GetRequired<IPartUriFeature>();
            destination.Features.Set<IPartUriFeature>(new CloningFeatures(existing));

            source.Save();

            foreach (var part in source.Parts)
            {
                destination.AddPart(part.OpenXmlPart, part.RelationshipId);
            }

            // Clone overloads without OpenSettings keep the source's current namespace prefix settings, including ones applied with
            // UseNamespacePrefixes after it was opened; overloads with OpenSettings use exactly those, applied when the clone was opened
            var namespacePrefixes = inheritNamespacePrefixes
                ? source.Features.Get<IXmlNamespacePrefixFeature>()
                : destination.Features.Get<IXmlNamespacePrefixFeature>();

            // the feature is not rebuilt from the source's OpenSettings, whose namespace prefix settings may have been changed since
            destination.SetOpenSettings(new(source.OpenSettings), namespacePrefixes);

            destination.Features.Set<IPartUriFeature>(existing);

            return destination;
        }
    }

    // a clone that keeps the source's namespace prefix settings takes them from its feature, so the ones in its OpenSettings, which
    // may have been changed since the source was opened, are neither validated nor applied
    private static OpenSettings ForClone(OpenSettings openSettings, bool inheritNamespacePrefixes)
        => inheritNamespacePrefixes && openSettings.NamespacePrefixes is not null ? new(openSettings) { NamespacePrefixes = null } : openSettings;

    internal static TPackage Reload<TPackage>(this TPackage openXmlPackage, bool? isEditable = default)
        where TPackage : OpenXmlPackage
    {
        if (openXmlPackage.Features.Get<IPackageFeature>() is { } package && package.Capabilities.HasFlagFast(PackageCapabilities.Reload))
        {
            if (isEditable.HasValue)
            {
                package.Reload(access: isEditable.Value ? FileAccess.ReadWrite : FileAccess.Read);
            }
            else
            {
                package.Reload();
            }
        }

        return openXmlPackage;
    }

    private sealed class CloningFeatures : IPartUriFeature
    {
        private readonly IPartUriFeature _other;

        public CloningFeatures(IPartUriFeature other)
        {
            _other = other;
        }

        Uri IPartUriFeature.CreatePartUri(string contentType, Uri parentUri, string targetPath, string targetName, string targetExt, bool forceUnique)
            => _other.CreatePartUri(contentType, parentUri, targetPath, targetName, targetExt, forceUnique: false);

        Uri IPartUriFeature.EnsureUniquePartUri(string contentType, Uri parentUri, Uri targetUri)
        {
            _other.ReserveUri(contentType, targetUri);
            return targetUri;
        }

        void IPartUriFeature.ReserveUri(string contentType, Uri partUri)
            => _other.ReserveUri(contentType, partUri);
    }
}
