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

using DocumentFormat.OpenXml.Packaging;
using Serialize.OpenXml.CodeGen;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Linq;

namespace DocxToSource.Wpf.Controls
{
    /// <summary>
    /// <see cref="System.Windows.Controls.TreeViewItem"/> derived class that contains information
    /// about a given <see cref="OpenXmlPart"/> object.
    /// </summary>
    public class OpenXmlPartTreeViewItem : OpenXmlTreeViewItem
    {
        #region Private Instance Fields

        /// <summary>
        /// Holds the <see cref="IdPartPair"/> object for this instance.
        /// </summary>
        private readonly IdPartPair part;

        #endregion

        #region Public Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenXmlPartTreeViewItem"/> class
        /// with the referenced <see cref="IdPartPair"/> object.
        /// </summary>
        /// <param name="p">
        /// The <see cref="IdPartPair"/> object that the new instance will hold.
        /// </param>
        /// <exception cref="ArgumentNullException">
        /// <paramref name="p"/> is <see langword="null"/>.
        /// </exception>
        public OpenXmlPartTreeViewItem(IdPartPair p)
            : base(p != null && p.OpenXmlPart != null && p.OpenXmlPart.Parts != null &&
                  (p.OpenXmlPart.Parts.Count() > 0 || p.OpenXmlPart.RootElement != null))
        {
            part = p ?? throw new ArgumentNullException(nameof(p));
        }

        #endregion

        #region Public Instance Methods

        /// <inheritdoc/>
        public override string BuildCodeDomTextDocument(CodeDomProvider provider)
        {
            return part.OpenXmlPart.GenerateSourceCode(provider);
        }

        #endregion

        #region Protected Instance Methods

        /// <inheritdoc/>
        protected override IEnumerable<OpenXmlTreeViewItem> GetOpenXmlSubtreeItems()
        {
            var result = new List<OpenXmlTreeViewItem>();

            result.AddRange(BuildOpenXmlPartTreeViewItems(part.OpenXmlPart.Parts));
            if (part.OpenXmlPart.RootElement != null)
            {
                result.AddRange(BuildOpenXmlElementTreeViewItems(part.OpenXmlPart.RootElement));
            }

            return result;
        }

        #endregion
    }
}
