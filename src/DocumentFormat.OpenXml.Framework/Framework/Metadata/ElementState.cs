// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace DocumentFormat.OpenXml.Framework.Metadata
{
    internal readonly struct ElementState
    {
        // Only the attribute values are stored as the attribute metadata is available from the element metadata,
        // which saves a field on every element. They are created separately from the metadata, as parents need their
        // metadata to create children regardless of whether they have any attributes.
        private readonly OpenXmlSimpleType?[]? _attributeData;

        public bool IsEmpty => Metadata is null;

        public bool HasAttributeData => _attributeData is not null;

        public ElementState(IElementMetadata metadata)
            : this(metadata, AttributeCollection.CreateData(metadata.Attributes))
        {
        }

        private ElementState(IElementMetadata metadata, OpenXmlSimpleType?[]? attributeData)
        {
            Metadata = metadata;
            _attributeData = attributeData;
        }

        /// <summary>
        /// Creates a state with only the metadata; <see cref="Attributes"/> is not available until created with <see cref="ElementState(IElementMetadata)"/>.
        /// </summary>
        public static ElementState MetadataOnly(IElementMetadata metadata) => new(metadata, null);

        /// <summary>
        /// Gets the attributes. Entries are nil if the state was created with <see cref="MetadataOnly(IElementMetadata)"/>.
        /// </summary>
        public AttributeCollection Attributes => new(Metadata.Attributes, _attributeData!);

        public IElementMetadata Metadata { get; }
    }
}
