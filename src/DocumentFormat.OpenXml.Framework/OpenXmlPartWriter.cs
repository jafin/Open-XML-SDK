// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using DocumentFormat.OpenXml.Features;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
#if FEATURE_ASYNC_SAX_XML
using DocumentFormat.OpenXml.Framework;
using System.Threading.Tasks;
#endif
using System.Xml;

namespace DocumentFormat.OpenXml
{
    /// <summary>
    /// Defines the OpenXmlPartWriter.
    /// </summary>
    public class OpenXmlPartWriter : OpenXmlWriter
    {
        private readonly XmlWriter _xmlWriter;
        private readonly IXmlNamespacePrefixFeature? _namespacePrefixes;

        // the default namespace in scope for each element started by this writer
        private readonly Stack<string> _defaultNamespaces = new();
        private bool _isLeafTextElementStart; // default is false

        // private Stack<OpenXmlElement> _elementStack;

        /// <summary>
        /// Initializes a new instance of the OpenXmlPartWriter.
        /// </summary>
        /// <param name="openXmlPart">The OpenXmlPart to be written to.</param>
        public OpenXmlPartWriter(OpenXmlPart openXmlPart)
            : this(openXmlPart, Encoding.UTF8)
        {
        }

        /// <summary>
        /// Initializes a new instance of the OpenXmlPartWriter.
        /// </summary>
        /// <param name="openXmlPart">The OpenXmlPart to be written to.</param>
        /// <param name="encoding">The encoding for the XML stream.</param>
        public OpenXmlPartWriter(OpenXmlPart openXmlPart, Encoding encoding)
        {
            if (openXmlPart is null)
            {
                throw new ArgumentNullException(nameof(openXmlPart));
            }

            if (encoding is null)
            {
                throw new ArgumentNullException(nameof(encoding));
            }

            var partStream = openXmlPart.GetStream(FileMode.Create);
            var settings = new XmlWriterSettings
            {
                CloseOutput = true,
                Encoding = encoding,
            };

            _xmlWriter = XmlWriter.Create(partStream, settings);
            _namespacePrefixes = openXmlPart.Features.Get<IXmlNamespacePrefixFeature>();
        }

        /// <summary>
        /// Initializes a new instance of the OpenXmlPartWriter.
        /// </summary>
        /// <param name="openXmlPart">The OpenXmlPart to be written to.</param>
        /// <param name="settings">The settings for the OpenXmlPartWriter.</param>
        public OpenXmlPartWriter(OpenXmlPart openXmlPart, OpenXmlPartWriterSettings settings)
        {
            if (openXmlPart is null)
            {
                throw new ArgumentNullException(nameof(openXmlPart));
            }

            if (settings is null)
            {
                throw new ArgumentNullException(nameof(settings));
            }

            var partStream = openXmlPart.GetStream(FileMode.Create);
            XmlWriterSettings xmlWriterSettings = new()
            {
                CloseOutput = true,
                Encoding = settings.Encoding,
#if FEATURE_ASYNC_SAX_XML
                Async = settings.Async,
#endif
            };

            _xmlWriter = XmlWriter.Create(partStream, xmlWriterSettings);
            _namespacePrefixes = settings.NamespacePrefixes is { } namespacePrefixes
                ? new XmlNamespacePrefixFeature(namespacePrefixes)
                : openXmlPart.Features.Get<IXmlNamespacePrefixFeature>();
        }

        /// <summary>
        /// Initializes a new instance of the OpenXmlPartWriter.
        /// </summary>
        /// <param name="partStream">The given part stream.</param>
        public OpenXmlPartWriter(Stream partStream)
            : this(partStream, Encoding.UTF8)
        {
        }

        /// <summary>
        /// Initializes a new instance of the OpenXmlPartWriter.
        /// </summary>
        /// <param name="partStream">The given part stream.</param>
        /// <param name="encoding">The encoding for the XML stream.</param>
        public OpenXmlPartWriter(Stream partStream, Encoding encoding)
        {
            if (partStream is null)
            {
                throw new ArgumentNullException(nameof(partStream));
            }

            if (encoding is null)
            {
                throw new ArgumentNullException(nameof(encoding));
            }

            var settings = new XmlWriterSettings
            {
                CloseOutput = false,
                Encoding = encoding,
            };

            _xmlWriter = XmlWriter.Create(partStream, settings);
        }

        /// <summary>
        /// Initializes a new instance of the OpenXmlPartWriter.
        /// </summary>
        /// <param name="partStream">The given part stream.</param>
        /// <param name="settings">The settings for the OpenXmlPartWriter.</param>
        public OpenXmlPartWriter(Stream partStream, OpenXmlPartWriterSettings settings)
        {
            if (partStream is null)
            {
                throw new ArgumentNullException(nameof(partStream));
            }

            if (settings is null)
            {
                throw new ArgumentNullException(nameof(settings));
            }

            XmlWriterSettings xmlWriterSettings = new()
            {
                CloseOutput = settings.CloseOutput,
                Encoding = settings.Encoding,
#if FEATURE_ASYNC_SAX_XML
                Async = settings.Async,
#endif
            };

            _xmlWriter = XmlWriter.Create(partStream, xmlWriterSettings);
            _namespacePrefixes = settings.NamespacePrefixes is { } namespacePrefixes ? new XmlNamespacePrefixFeature(namespacePrefixes) : null;
        }

        #region public OpenXmlWriter methods

        /// <summary>
        /// Writes the XML declaration with the version "1.0".
        /// </summary>
        public override void WriteStartDocument()
        {
            ThrowIfObjectDisposed();

            _xmlWriter.WriteStartDocument();
        }

        /// <summary>
        /// Writes the XML declaration with the version "1.0" and the standalone attribute.
        /// </summary>
        /// <param name="standalone">If true, it writes "standalone=yes"; if false, it writes "standalone=no". </param>
        public override void WriteStartDocument(bool standalone)
        {
            ThrowIfObjectDisposed();

            _xmlWriter.WriteStartDocument(standalone);
        }

        /// <summary>
        /// Writes out a start element tag of the current element of the OpenXmlReader. And write all the attributes of the element.
        /// </summary>
        /// <param name="elementReader">The OpenXmlReader to read from. </param>
        public override void WriteStartElement(OpenXmlReader elementReader)
        {
            if (elementReader is null)
            {
                throw new ArgumentNullException(nameof(elementReader));
            }

            WriteStartElement(elementReader, elementReader.Attributes, elementReader.NamespaceDeclarations);
        }

        /// <summary>
        /// Writes out a start element tag of the current element of the OpenXmlReader. And write the attributes in attributes.
        /// </summary>
        /// <param name="elementReader">The OpenXmlReader to read from. </param>
        /// <param name="attributes">The attributes to be written, can be null if no attributes.</param>
        public override void WriteStartElement(OpenXmlReader elementReader, IEnumerable<OpenXmlAttribute> attributes)
        {
            if (elementReader is null)
            {
                throw new ArgumentNullException(nameof(elementReader));
            }

            WriteStartElement(elementReader, attributes, elementReader.NamespaceDeclarations);
        }

        /// <summary>
        /// Writes out a start element tag of the current element of the OpenXmlReader. And write the attributes in attributes.
        /// </summary>
        /// <param name="elementReader">The OpenXmlReader to read from. </param>
        /// <param name="attributes">The attributes to be written, can be null if no attributes.</param>
        /// <param name="namespaceDeclarations">The namespace declarations to be written, can be null if no namespace declarations.</param>
        public override void WriteStartElement(OpenXmlReader elementReader, IEnumerable<OpenXmlAttribute> attributes, IEnumerable<KeyValuePair<string, string>> namespaceDeclarations)
        {
            if (elementReader is null)
            {
                throw new ArgumentNullException(nameof(elementReader));
            }

            if (elementReader.IsEndElement)
            {
                throw new ArgumentOutOfRangeException(nameof(elementReader));
            }

            if (elementReader.IsMiscNode)
            {
                // OpenXmlMiscNode should be written by WriteElement( );
                throw new ArgumentOutOfRangeException(nameof(elementReader));
            }

            ThrowIfObjectDisposed();

            WriteStartTag(elementReader.Prefix, elementReader.LocalName, elementReader.NamespaceUri, attributes, namespaceDeclarations);

            if (attributes is not null)
            {
                // write attributes
                foreach (var attribute in attributes)
                {
                    _xmlWriter.WriteAttributeString(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, attribute.Value);
                }
            }

            if (elementReader.ElementType.IsSubclassOf(typeof(OpenXmlLeafTextElement)))
            {
                _isLeafTextElementStart = true;
            }
            else
            {
                _isLeafTextElementStart = false;
            }
        }

        /// <summary>
        /// Writes out a start tag of the element and all the attributes of the element.
        /// </summary>
        /// <param name="elementObject">The OpenXmlElement object to be written.</param>
        public override void WriteStartElement(OpenXmlElement elementObject)
        {
            if (elementObject is null)
            {
                throw new ArgumentNullException(nameof(elementObject));
            }

            if (elementObject is OpenXmlMiscNode)
            {
                throw new ArgumentOutOfRangeException(nameof(elementObject));
            }

            ThrowIfObjectDisposed();

            WriteStartTag(elementObject.Prefix, elementObject.LocalName, elementObject.NamespaceUri, attributes: null, namespaceDeclarations: null);

            if (elementObject.HasAttributes)
            {
                // write attributes
                foreach (var attribute in elementObject.GetAttributes())
                {
                    _xmlWriter.WriteAttributeString(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, attribute.Value);
                }
            }

            if (elementObject is OpenXmlLeafTextElement)
            {
                _isLeafTextElementStart = true;
            }
            else
            {
                _isLeafTextElementStart = false;
            }
        }

        /// <summary>
        /// Writes out a start tag of the element. And write the attributes in attributes. The attributes of the element will be omitted.
        /// </summary>
        /// <param name="elementObject">The OpenXmlElement object to be written.</param>
        /// <param name="attributes">The attributes to be written.</param>
        public override void WriteStartElement(OpenXmlElement elementObject, IEnumerable<OpenXmlAttribute> attributes)
        {
            if (elementObject is null)
            {
                throw new ArgumentNullException(nameof(elementObject));
            }

            WriteStartElement(elementObject, attributes, elementObject.NamespaceDeclarations);
        }

        /// <summary>
        /// Writes out a start tag of the element. And write the attributes in attributes. The attributes of the element will be omitted.
        /// </summary>
        /// <param name="elementObject">The OpenXmlElement object to be written.</param>
        /// <param name="attributes">The attributes to be written.</param>
        /// <param name="namespaceDeclarations">The namespace declarations to be written, can be null if no namespace declarations.</param>
        public override void WriteStartElement(OpenXmlElement elementObject, IEnumerable<OpenXmlAttribute> attributes, IEnumerable<KeyValuePair<string, string>> namespaceDeclarations)
        {
            if (elementObject is null)
            {
                throw new ArgumentNullException(nameof(elementObject));
            }

            if (elementObject is OpenXmlMiscNode)
            {
                throw new ArgumentOutOfRangeException(nameof(elementObject));
            }

            ThrowIfObjectDisposed();

            WriteStartTag(elementObject.Prefix, elementObject.LocalName, elementObject.NamespaceUri, attributes, namespaceDeclarations);

            if (attributes is not null)
            {
                // write attributes
                foreach (var attribute in attributes)
                {
                    _xmlWriter.WriteAttributeString(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, attribute.Value);
                }
            }

            if (elementObject is OpenXmlLeafTextElement)
            {
                _isLeafTextElementStart = true;
            }
            else
            {
                _isLeafTextElementStart = false;
            }
        }

        /// <summary>
        /// Closes one element.
        /// </summary>
        public override void WriteEndElement()
        {
            ThrowIfObjectDisposed();

            _xmlWriter.WriteEndElement();
            PopDefaultNamespace();

            _isLeafTextElementStart = false;
        }

        /// <summary>
        /// Writes the given text content.
        /// </summary>
        /// <param name="text">The text to be written. </param>
        public override void WriteString(string text)
        {
            ThrowIfObjectDisposed();

            if (_isLeafTextElementStart)
            {
                _xmlWriter.WriteString(text);
            }
            else
            {
                throw new InvalidOperationException(ExceptionMessages.InvalidWriteStringCall);
            }

            // can continue WriteString(), so don't set _isLeafTextElementStart to false.
        }

        /// <summary>
        /// Write the OpenXmlElement to the writer.
        /// </summary>
        /// <param name="elementObject">The OpenXmlElement object to be written.</param>
        public override void WriteElement(OpenXmlElement elementObject)
        {
            if (elementObject is null)
            {
                throw new ArgumentNullException(nameof(elementObject));
            }

            ThrowIfObjectDisposed();

            if (CurrentDefaultNamespace.Length > 0)
            {
                // track the default namespace so the subtree is written without prefixes where possible
                using var writer = new XmlDOMTextWriter(_xmlWriter, CurrentDefaultNamespace);
                elementObject.WriteTo(writer);
            }
            else
            {
                elementObject.WriteTo(_xmlWriter);
            }

            _isLeafTextElementStart = false;
        }

        /// <summary>
        /// Close the writer.
        /// </summary>
        public override void Close()
        {
            if (_xmlWriter is not null)
            {
                _xmlWriter.Close();
            }

            _isLeafTextElementStart = false;
        }

        private static IOpenXmlNamespaceResolver Resolver => FeatureCollection.Default.GetNamespaceResolver();

        private string CurrentDefaultNamespace => _defaultNamespaces.Count == 0 ? string.Empty : _defaultNamespaces.Peek();

        private void PopDefaultNamespace()
        {
            if (_defaultNamespaces.Count > 0)
            {
                _defaultNamespaces.Pop();
            }
        }

        private void WriteStartTag(string? prefix, string localName, string namespaceUri, IEnumerable<OpenXmlAttribute>? attributes, IEnumerable<KeyValuePair<string, string>>? namespaceDeclarations)
        {
            var rootPrefix = StartElementScope(ref prefix, namespaceUri, attributes, namespaceDeclarations, out var skipLocalDefault);

            _xmlWriter.WriteStartElement(prefix, localName, namespaceUri);

            if (namespaceDeclarations is not null)
            {
                foreach (var item in namespaceDeclarations)
                {
                    if (!skipLocalDefault || item.Key.Length != 0)
                    {
                        _xmlWriter.WriteNamespaceDeclaration(item.Key, item.Value);
                    }
                }
            }

            if (rootPrefix is not null)
            {
                _xmlWriter.WriteNamespaceDeclaration(rootPrefix, namespaceUri);
            }
        }

        /// <summary>
        /// Chooses the prefix for an element's start tag and tracks the default namespace in scope for its content.
        /// </summary>
        /// <param name="prefix">The prefix of the element, updated to the prefix to write.</param>
        /// <param name="namespaceUri">The namespace of the element.</param>
        /// <param name="attributes">The attributes to be written, which may declare the default namespace.</param>
        /// <param name="namespaceDeclarations">The namespace declarations to be written.</param>
        /// <param name="skipLocalDefault">Whether a default namespace declaration in <paramref name="namespaceDeclarations"/> must be skipped.</param>
        /// <returns>The prefix to declare for the element's namespace when the root is written in the default namespace.</returns>
        private string? StartElementScope(ref string? prefix, string namespaceUri, IEnumerable<OpenXmlAttribute>? attributes, IEnumerable<KeyValuePair<string, string>>? namespaceDeclarations, out bool skipLocalDefault)
        {
            string? localDefault = null;
            var hasPrefixedDeclaration = false;

            if (namespaceDeclarations is not null)
            {
                foreach (var item in namespaceDeclarations)
                {
                    if (item.Key.Length == 0)
                    {
                        localDefault = item.Value;
                    }
                    else if (item.Value == namespaceUri)
                    {
                        hasPrefixedDeclaration = true;
                    }
                }
            }

            if (attributes is not null)
            {
                foreach (var attribute in attributes)
                {
                    if (attribute.Prefix.Length == 0 && attribute.LocalName == OpenXmlElementContext.XmlnsPrefix)
                    {
                        localDefault = attribute.Value;
                    }
                }
            }

            skipLocalDefault = false;
            string? rootPrefix = null;

            if (_defaultNamespaces.Count == 0 && _namespacePrefixes.IsDefaultNamespaceForRoot(namespaceUri))
            {
                prefix = string.Empty;
                skipLocalDefault = localDefault is not null && localDefault != namespaceUri;
                localDefault = namespaceUri;

                // the content cannot be inspected in advance, so declare the prefix in case an attribute needs it
                if (!hasPrefixedDeclaration && Resolver.LookupPrefix(namespaceUri) is { Length: > 0 } builtInPrefix)
                {
                    rootPrefix = builtInPrefix;
                }
            }
            else if (localDefault is not null)
            {
                if (localDefault == namespaceUri)
                {
                    prefix = string.Empty;
                }
                else if (string.IsNullOrEmpty(prefix))
                {
                    prefix = Resolver.LookupPrefix(namespaceUri);
                }
            }
            else if (namespaceUri.Length > 0 && CurrentDefaultNamespace == namespaceUri)
            {
                prefix = string.Empty;
            }

            // a null prefix lets the writer pick one, so resolve it here to know which default namespace is in scope
            prefix ??= _xmlWriter.LookupPrefix(namespaceUri) ?? string.Empty;

            _defaultNamespaces.Push(localDefault ?? (prefix.Length == 0 ? namespaceUri : CurrentDefaultNamespace));

            return rootPrefix;
        }

#if FEATURE_ASYNC_SAX_XML
        private async Task WriteStartTagAsync(string? prefix, string localName, string namespaceUri, IEnumerable<OpenXmlAttribute>? attributes, IEnumerable<KeyValuePair<string, string>>? namespaceDeclarations)
        {
            var rootPrefix = StartElementScope(ref prefix, namespaceUri, attributes, namespaceDeclarations, out var skipLocalDefault);

            await _xmlWriter.WriteStartElementAsync(prefix, localName, namespaceUri).ConfigureAwait(true);

            if (namespaceDeclarations is not null)
            {
                foreach (var item in namespaceDeclarations)
                {
                    if (!skipLocalDefault || item.Key.Length != 0)
                    {
                        await _xmlWriter.WriteNamespaceDeclarationAsync(item.Key, item.Value).ConfigureAwait(true);
                    }
                }
            }

            if (rootPrefix is not null)
            {
                await _xmlWriter.WriteNamespaceDeclarationAsync(rootPrefix, namespaceUri).ConfigureAwait(true);
            }
        }
#endif

        #endregion

        // Async Methods
#if FEATURE_ASYNC_SAX_XML
        /// <summary>
        /// Asynchronously writes the XML declaration with the version "1.0".
        /// </summary>
        public override Task WriteStartDocumentAsync()
        {
            ThrowIfObjectDisposed();

            return _xmlWriter.WriteStartDocumentAsync();
        }

        /// <summary>
        /// Asynchronously writes the XML declaration with the version "1.0" and the standalone attribute.
        /// </summary>
        /// <param name="standalone">If true, it writes "standalone=yes"; if false, it writes "standalone=no". </param>
        public override Task WriteStartDocumentAsync(bool standalone)
        {
            ThrowIfObjectDisposed();

            return _xmlWriter.WriteStartDocumentAsync(standalone);
        }

        /// <summary>
        /// Asynchronously writes out a start tag of the element and all the attributes of the element.
        /// </summary>
        /// <param name="elementObject">The OpenXmlElement object to be written.</param>
        public async override Task WriteStartElementAsync(OpenXmlElement elementObject)
        {
            if (elementObject is null)
            {
                throw new ArgumentNullException(nameof(elementObject));
            }

            if (elementObject is OpenXmlMiscNode)
            {
                throw new ArgumentOutOfRangeException(nameof(elementObject));
            }

            ThrowIfObjectDisposed();

            await WriteStartTagAsync(elementObject.Prefix, elementObject.LocalName, elementObject.NamespaceUri, attributes: null, namespaceDeclarations: null).ConfigureAwait(true);

            if (elementObject.HasAttributes)
            {
                // write attributes
                foreach (var attribute in elementObject.GetAttributes())
                {
                    await _xmlWriter.WriteAttributeStringAsync(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, attribute.Value).ConfigureAwait(true);
                }
            }

            if (elementObject is OpenXmlLeafTextElement)
            {
                _isLeafTextElementStart = true;
            }
            else
            {
                _isLeafTextElementStart = false;
            }
        }

        /// <summary>
        /// Asynchronously writes out a start tag of the element. And write the attributes in attributes. The attributes of the element will be omitted.
        /// </summary>
        /// <param name="elementObject">The OpenXmlElement object to be written.</param>
        /// <param name="attributes">The attributes to be written.</param>
        public override Task WriteStartElementAsync(OpenXmlElement elementObject, IEnumerable<OpenXmlAttribute> attributes)
        {
            if (elementObject is null)
            {
                throw new ArgumentNullException(nameof(elementObject));
            }

            return WriteStartElementAsync(elementObject, attributes, elementObject.NamespaceDeclarations);
        }

        /// <summary>
        /// Asynchronously writes out a start tag of the element. And write the attributes in attributes. The attributes of the element will be omitted.
        /// </summary>
        /// <param name="elementObject">The OpenXmlElement object to be written.</param>
        /// <param name="attributes">The attributes to be written.</param>
        /// <param name="namespaceDeclarations">The namespace declarations to be written, can be null if no namespace declarations.</param>
        public async override Task WriteStartElementAsync(OpenXmlElement elementObject, IEnumerable<OpenXmlAttribute> attributes, IEnumerable<KeyValuePair<string, string>> namespaceDeclarations)
        {
            if (elementObject is null)
            {
                throw new ArgumentNullException(nameof(elementObject));
            }

            if (elementObject is OpenXmlMiscNode)
            {
                throw new ArgumentOutOfRangeException(nameof(elementObject));
            }

            ThrowIfObjectDisposed();

            await WriteStartTagAsync(elementObject.Prefix, elementObject.LocalName, elementObject.NamespaceUri, attributes, namespaceDeclarations).ConfigureAwait(true);

            if (attributes is not null)
            {
                // write attributes
                foreach (var attribute in attributes)
                {
                    await _xmlWriter.WriteAttributeStringAsync(attribute.Prefix, attribute.LocalName, attribute.NamespaceUri, attribute.Value).ConfigureAwait(true);
                }
            }

            if (elementObject is OpenXmlLeafTextElement)
            {
                _isLeafTextElementStart = true;
            }
            else
            {
                _isLeafTextElementStart = false;
            }
        }

        /// <summary>
        /// Asynchronously closes one element.
        /// </summary>
        public override Task WriteEndElementAsync()
        {
            ThrowIfObjectDisposed();

            _isLeafTextElementStart = false;
            PopDefaultNamespace();

            return _xmlWriter.WriteEndElementAsync();
        }

        /// <summary>
        /// Asynchronously writes the OpenXmlElement to the writer.
        /// </summary>
        /// <param name="elementObject">The OpenXmlElement object to be written.</param>
        public async override Task WriteElementAsync(OpenXmlElement elementObject)
        {
            if (elementObject is null)
            {
                throw new ArgumentNullException(nameof(elementObject));
            }

            ThrowIfObjectDisposed();

            await WriteStartElementAsync(elementObject).ConfigureAwait(true);

            await WriteEndElementAsync().ConfigureAwait(true);

            _isLeafTextElementStart = false;
        }

        /// <summary>
        /// Asynchronously writes the given text content.
        /// </summary>
        /// <param name="text">The text to be written. </param>
        public override Task WriteStringAsync(string text)
        {
            ThrowIfObjectDisposed();

            if (_isLeafTextElementStart)
            {
                return _xmlWriter.WriteStringAsync(text);
            }
            else
            {
                throw new InvalidOperationException(ExceptionMessages.InvalidWriteStringCall);
            }

            // can continue WriteStringAsync(), so don't set _isLeafTextElementStart to false.
        }
#endif
    }
}
