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
using System.Collections.ObjectModel;

namespace Serialize.OpenXml.CodeGen
{
    /// <summary>
    /// Collection class used to organize all of the <see cref="OpenXmlPartBluePrint"/> objects
    /// for a single request.
    /// </summary>
    public sealed class OpenXmlPartBluePrintCollection : KeyedCollection<Uri, OpenXmlPartBluePrint>
    {
        #region Internal Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenXmlPartBluePrintCollection"/> class
        /// that is empty.
        /// </summary>
        internal OpenXmlPartBluePrintCollection() : base(new UriEqualityComparer()) { }

        #endregion

        #region Public Instance Methods

        /// <summary>
        /// Gets the <see cref="OpenXmlPartBluePrint"/> object associated with the specified
        /// <see cref="Uri"/>.
        /// </summary>
        /// <param name="uri">
        /// The <see cref="Uri"/> of the <see cref="OpenXmlPartBluePrint"/> object to get.
        /// </param>
        /// <param name="bp">
        /// When the method returns, contains the <see cref="OpenXmlPartBluePrint"/> object
        /// associated with <paramref name="uri"/>, if found; otherwise, the default value
        /// of <see cref="OpenXmlPartBluePrint"/> is returned (<see langword="null"/>).
        /// This parameter is passed uninitialized.
        /// </param>
        /// <returns>
        /// <see langword="true"/> if collection contains an <see cref="OpenXmlPartBluePrint"/>
        /// object with the specified <paramref name="uri"/>; otherwise, <see langword="false"/>.
        /// </returns>
        public bool TryGetValue(Uri uri, out OpenXmlPartBluePrint bp)
        {
            if (uri == null || !Contains(uri))
            {
                bp = null;
                return false;
            }
            bp = this[uri];
            return true;
        }

        #endregion

        #region Protected Instance Methods

        /// <inheritdoc/>
        protected override Uri GetKeyForItem(OpenXmlPartBluePrint item) => item.Uri;

        #endregion
    }
}