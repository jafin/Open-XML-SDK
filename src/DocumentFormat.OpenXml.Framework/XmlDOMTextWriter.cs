// Copyright (c) Microsoft. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;

namespace DocumentFormat.OpenXml
{
    internal class XmlDOMTextWriter : XmlWriter
    {
        private readonly XmlWriter _writer;
        private readonly bool _ownsWriter = true;

        // the number of scopes pushed before anything is written, i.e. by a wrapper that starts inside an element
        private readonly int _topLevelDepth;

        // XmlWriter.LookupPrefix prefers a prefixed binding over the default namespace when both map to the same
        // uri, so track the default namespace in scope ourselves to allow elements to be written without a prefix.
        private readonly Stack<string> _defaultNamespaces = new();
        private string? _pendingDefaultNamespace;

        public XmlDOMTextWriter(Stream stream)
        {
            _writer = Create(stream);
        }

        public XmlDOMTextWriter(Stream stream, XmlWriterSettings settings)
        {
            _writer = Create(stream, settings);
        }

        public XmlDOMTextWriter(TextWriter w)
        {
            var xwSettings = new XmlWriterSettings
            {
                Encoding = w.Encoding,
                OmitXmlDeclaration = true,
                ConformanceLevel = ConformanceLevel.Fragment,
            };

            _writer = Create(w, xwSettings);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="XmlDOMTextWriter"/> class that writes to an existing writer, which it
        /// does not close, starting inside an element whose default namespace is <paramref name="defaultNamespace"/>.
        /// </summary>
        public XmlDOMTextWriter(XmlWriter writer, string defaultNamespace)
        {
            _writer = writer;
            _ownsWriter = false;
            _defaultNamespaces.Push(defaultNamespace);
            _topLevelDepth = 1;
        }

        /// <summary>
        /// Gets the default namespace in scope for the element currently being written.
        /// </summary>
        internal string DefaultNamespace => _defaultNamespaces.Count == 0 ? string.Empty : _defaultNamespaces.Peek();

        /// <summary>
        /// Gets or sets a value that, when set, decides whether a part root element written to this writer uses the default namespace,
        /// instead of the settings of the part the root belongs to.
        /// </summary>
        internal bool? UseDefaultNamespaceForRoot { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether a part root element is writing its attributes while written in the default namespace.
        /// </summary>
        internal bool IsWritingRootInDefaultNamespace { get; set; }

        /// <summary>
        /// Gets a value indicating whether no element started on this writer is open, so the next element is the first one of the write.
        /// </summary>
        internal bool IsAtTopLevel => _defaultNamespaces.Count <= _topLevelDepth;

        public override WriteState WriteState => _writer.WriteState;

        public override void Flush() => _writer.Flush();

        public override string? LookupPrefix(string ns) => _writer.LookupPrefix(ns);

        public override void WriteBase64(byte[] buffer, int index, int count) => _writer.WriteBase64(buffer, index, count);

        public override void WriteCData(string? text) => _writer.WriteCData(text);

        public override void WriteCharEntity(char ch) => _writer.WriteCharEntity(ch);

        public override void WriteChars(char[] buffer, int index, int count) => _writer.WriteChars(buffer, index, count);

        public override void WriteComment(string? text) => _writer.WriteComment(text);

        public override void WriteDocType(string name, string? pubid, string? sysid, string? subset) => _writer.WriteDocType(name, pubid, sysid, subset);

        public override void WriteEndAttribute()
        {
            _writer.WriteEndAttribute();

            if (_pendingDefaultNamespace is not null)
            {
                _defaultNamespaces.Pop();
                _defaultNamespaces.Push(_pendingDefaultNamespace);
                _pendingDefaultNamespace = null;
            }
        }

        public override void WriteEndDocument() => _writer.WriteEndDocument();

        public override void WriteEndElement()
        {
            _writer.WriteEndElement();
            PopDefaultNamespace();
        }

        public override void WriteEntityRef(string name) => _writer.WriteEntityRef(name);

        public override void WriteFullEndElement()
        {
            _writer.WriteFullEndElement();
            PopDefaultNamespace();
        }

        private void PopDefaultNamespace()
        {
            if (_defaultNamespaces.Count > 0)
            {
                _defaultNamespaces.Pop();
            }
        }

        public override void WriteProcessingInstruction(string name, string? text) => _writer.WriteProcessingInstruction(name, text);

        public override void WriteRaw(string data) => _writer.WriteRaw(data);

        public override void WriteRaw(char[] buffer, int index, int count) => _writer.WriteRaw(buffer, index, count);

        public override void WriteStartAttribute(string? prefix, string localName, string? ns)
        {
            if (string.IsNullOrEmpty(localName))
            {
                throw new ArgumentNullException(nameof(localName));
            }

            if (prefix is null)
            {
                prefix = string.Empty;
            }

            if (ns is null)
            {
                ns = string.Empty;
            }

            if ((ns.Length == 0) && (prefix.Length != 0))
            {
                prefix = string.Empty;
            }

            _writer.WriteStartAttribute(prefix, localName, ns);

            if (prefix.Length == 0 && localName == OpenXmlElementContext.XmlnsPrefix && _defaultNamespaces.Count > 0)
            {
                _pendingDefaultNamespace = string.Empty;
            }
        }

        public override void WriteStartDocument() => _writer.WriteStartDocument();

        public override void WriteStartDocument(bool standalone) => _writer.WriteStartDocument(standalone);

        public override void WriteStartElement(string? prefix, string localName, string? ns)
        {
            if (string.IsNullOrEmpty(localName))
            {
                throw new ArgumentNullException(nameof(localName));
            }

            if (prefix is null)
            {
                prefix = string.Empty;
            }

            if (ns is null)
            {
                ns = string.Empty;
            }

            if ((ns.Length == 0) && (prefix.Length != 0))
            {
                prefix = string.Empty;
            }

            _writer.WriteStartElement(prefix, localName, ns);
            _defaultNamespaces.Push(prefix.Length == 0 ? ns : DefaultNamespace);
        }

        public override void WriteString(string? text)
        {
            if (!string.IsNullOrEmpty(text))
            {
                _writer.WriteString(text);

                if (_pendingDefaultNamespace is not null)
                {
                    _pendingDefaultNamespace += text;
                }
            }
        }

        public override void WriteSurrogateCharEntity(char lowChar, char highChar) => _writer.WriteSurrogateCharEntity(lowChar, highChar);

        public override void WriteWhitespace(string? ws) => _writer.WriteWhitespace(ws);

        public override XmlWriterSettings? Settings => _writer.Settings;

        public override string? XmlLang => _writer.XmlLang;

        public override XmlSpace XmlSpace => _writer.XmlSpace;

        public override void Close()
        {
            if (_ownsWriter)
            {
                _writer.Close();
            }
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);

            if (disposing && _ownsWriter)
            {
#if NET35 || NET40
                ((IDisposable)_writer).Dispose();
#else
                _writer.Dispose();
#endif
            }
        }
    }
}
