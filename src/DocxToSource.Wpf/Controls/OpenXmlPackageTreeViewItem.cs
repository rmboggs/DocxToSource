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
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Linq;

namespace DocxToSource.Wpf.Controls
{
    /// <summary>
    /// <see cref="OpenXmlTreeViewItem"/> derived class that contains information
    /// about a given <see cref="OpenXmlPackage"/> object.
    /// </summary>
    public class OpenXmlPackageTreeViewItem : OpenXmlTreeViewItem
    {
        #region Private Instance Fields

        /// <summary>
        /// Holds the OpenXmlPackage that is currently open.
        /// </summary>
        private readonly OpenXmlPackage _package;

        #endregion

        #region Public Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenXmlPackageTreeViewItem"/> class
        /// that is empty.
        /// </summary>
        public OpenXmlPackageTreeViewItem(OpenXmlPackage pkg)
            : base(pkg.Parts != null && pkg.Parts.Count() > 0) { _package = pkg; }

        #endregion

        #region Public Instance Methods

        /// <inheritdoc/>
        public override string BuildCodeDomTextDocument(CodeDomProvider provider)
        {
            return _package.GenerateSourceCode(provider);
        }

        #endregion

        #region Protected Instance Methods

        /// <inheritdoc/>
        protected override IEnumerable<OpenXmlTreeViewItem> GetOpenXmlSubtreeItems()
        {
            if (_package.Parts is null) return null;
            return BuildOpenXmlPartTreeViewItems(_package.Parts);
        }

        #endregion
    }
}
