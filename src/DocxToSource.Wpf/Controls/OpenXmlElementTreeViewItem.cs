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

using DocumentFormat.OpenXml;
using Serialize.OpenXml.CodeGen;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;

namespace DocxToSource.Wpf.Controls
{
    /// <summary>
    /// <see cref="OpenXmlTreeViewItem"/> derived class that contains information
    /// about a given <see cref="OpenXmlElement"/> object.
    /// </summary>
    public class OpenXmlElementTreeViewItem : OpenXmlTreeViewItem
    {
        #region Private Instance Fields

        /// <summary>
        /// Holds the <see cref="OpenXmlElement"/> object for this instance.
        /// </summary>
        private readonly OpenXmlElement element;

        #endregion

        #region Public Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenXmlElementTreeViewItem"/> class
        /// with the referenced <see cref="OpenXmlElement"/> object.
        /// </summary>
        /// <param name="e">
        /// The <see cref="OpenXmlElement"/> object that the new instance will hold.
        /// </param>
        /// <exception cref="ArgumentNullException">
        /// <paramref name="e"/> is <see langword="null"/>.
        /// </exception>
        public OpenXmlElementTreeViewItem(OpenXmlElement e) : base(e != null && e.HasChildren)
        {
            if (e is null) throw new ArgumentNullException(nameof(e));
            element = e;
        }

        #endregion

        #region Public Instance Methods

        /// <inheritdoc/>
        public override string BuildCodeDomTextDocument(CodeDomProvider provider)
        {
            return element.GenerateSourceCode(provider);
        }

        /// <inheritdoc/>
        public override string BuildXmlTextDocument()
        {
            var sb = new StringBuilder();

            using (var writer = new StringWriter(sb))
            {
                using (var xTarget = new XmlTextWriter(writer))
                {

                    var xDoc = new XmlDocument();
                    xDoc.LoadXml(element.OuterXml);

                    xTarget.Formatting = Formatting.Indented;
                    xTarget.Indentation = 2;

                    xDoc.Normalize();
                    xDoc.PreserveWhitespace = true;
                    xDoc.WriteContentTo(xTarget);

                    xTarget.Flush();
                    xTarget.Close();
                }
            }

            return sb.ToString();
        }

        #endregion

        #region Protected Instance Methods

        /// <inheritdoc/>
        protected override IEnumerable<OpenXmlTreeViewItem> GetOpenXmlSubtreeItems()
        {
            return BuildOpenXmlElementTreeViewItems(element.Elements());
        }

        #endregion
    }
}
