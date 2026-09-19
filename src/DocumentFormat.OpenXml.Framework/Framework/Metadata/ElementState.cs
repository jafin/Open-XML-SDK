// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

namespace DocumentFormat.OpenXml.Framework.Metadata
{
    internal readonly struct ElementState
    {
        // Only the attribute values are stored as the attribute metadata is available from the element metadata,
        // which saves a field on every element.
        private readonly OpenXmlSimpleType?[] _attributeData;

        public bool IsEmpty => _attributeData is null;

        public ElementState(IElementMetadata metadata)
        {
            _attributeData = AttributeCollection.CreateData(metadata.Attributes);
            Metadata = metadata;
        }

        public AttributeCollection Attributes => new(Metadata.Attributes, _attributeData);

        public IElementMetadata Metadata { get; }
    }
}
