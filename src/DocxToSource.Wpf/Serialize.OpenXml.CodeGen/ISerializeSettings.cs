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
using DocumentFormat.OpenXml;

namespace Serialize.OpenXml.CodeGen
{
    /// <summary>
    /// Defines objects containing settings to use when generating
    /// source code for OpenXml SDK based files.
    /// </summary>
    public interface ISerializeSettings
    {
        #region Properties

        /// <summary>
        /// Gets the lookup collection of <see cref="IOpenXmlHandler"/> objects to use
        /// for custom code generations for specific OpenXml SDK objects.
        /// </summary>
        /// <remarks>
        /// If this collection is <see langword="null"/> or if the current OpenXml SDK
        /// object type being analyzed does not exist, then the default code generation
        /// process will be used.
        /// </remarks>
        IReadOnlyDictionary<Type, IOpenXmlHandler> Handlers { get; }

        /// <summary>
        /// Gets all of the <see cref="XmlNodeType"/> setting values of the
        /// <see cref="OpenXmlMiscNode"/> objects to ignore.
        /// </summary>
        /// <remarks>
        /// If this collection is <see langword="null"/> or if the <see cref="XmlNodeType"/>
        /// value of a <see cref="OpenXmlMiscNode"/> object does not exist in this collection,
        /// then the <see cref="OpenXmlMiscNode"/> object will be processed as a normal
        /// <see cref="OpenXmlElement"/> object.
        /// </remarks>
        IReadOnlyList<XmlNodeType> IgnoreMiscNodeTypes { get; }

        /// <summary>
        /// Indicates whether or not to ignore <see cref="OpenXmlUnknownElement"/>
        /// objects when generating source code.
        /// </summary>
        bool IgnoreUnknownElements { get; }

        /// <summary>
        /// Gets the <see cref="NamespaceAliasOptions"/> object to use during the code
        /// generation process.
        /// </summary>s
        NamespaceAliasOptions NamespaceAliasOptions { get; }

        #endregion
    }
}