/* MIT License

Copyright (c) 2020 Ryan Boggs

Permission is hereby granted, free of charge, to any person obtaining a copy of this
software and associated documentation files (the "Software"), to deal in the Software
without restriction, including without limitation the rights to use, copy, modify,
merge, publish, distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to the following
conditions:

The above copyright notice and this permission notice shall be included in all copies
or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
DEALINGS IN THE SOFTWARE.
*/

using System;
using System.Collections.Generic;
using System.Xml;

namespace Serialize.OpenXml.CodeGen
{
    /// <summary>
    /// A default implementation of the <see cref="ISerializeSettings"/> interface
    /// to use when custom <see cref="ISerializeSettings"/> object is not provided.
    /// </summary>
    internal sealed class DefaultSerializeSettings : ISerializeSettings
    {
        #region Public Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="DefaultSerializeSettings"/>
        /// class that is empty.
        /// </summary>
        public DefaultSerializeSettings() : this(NamespaceAliasOptions.Default) {}

        /// <summary>
        /// Initializes a new instance of the <see cref="DefaultSerializeSettings"/>
        /// class with the requested <see cref="NamespaceAliasOptions"/> object.
        /// </summary>
        /// <param name="nsAliasOpts"></param>
        public DefaultSerializeSettings(NamespaceAliasOptions nsAliasOpts)
        {
            NamespaceAliasOptions = nsAliasOpts
                ?? throw new ArgumentNullException(nameof(nsAliasOpts));
        }

        #endregion

        #region Public Instance Properties

        /// <inheritdoc/>
        public IReadOnlyDictionary<Type, IOpenXmlHandler> Handlers => null;

        /// <inheritdoc/>
        public IReadOnlyList<XmlNodeType> IgnoreMiscNodeTypes => null;

        /// <inheritdoc/>
        public bool IgnoreUnknownElements => false;

        /// <inheritdoc/>
        public NamespaceAliasOptions NamespaceAliasOptions
        {
            get;
            private set;
        }

        #endregion
    }
}