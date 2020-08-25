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
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Controls;

namespace DocxToSource.Wpf.Controls
{
    /// <summary>
    /// Base <see cref="TreeViewItem"/> class to use for all openxml sdk objects of a 
    /// selected document.
    /// </summary>
    public abstract class OpenXmlTreeViewItem : TreeViewItem
    {
        #region Protected Static Fields

        /// <summary>
        /// Holds the code generation options to use when building the source code text.
        /// </summary>
        protected readonly CodeGeneratorOptions Cgo = new CodeGeneratorOptions()
        {
            BracingStyle = "C"
        };

        #endregion

        #region Private Instance Fields

        /// <summary>
        /// Indicates whether or not the child tree view items have been loaded.
        /// </summary>
        private bool _subtreeIsLoaded = false;

        #endregion

        #region Protected Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenXmlTreeViewItem"/> class that is
        /// empty.
        /// </summary>
        /// <param name="hasChildren">
        /// Indicates whether or not the new object is expected to have sub tree items.
        /// </param>
        protected OpenXmlTreeViewItem(bool hasChildren) : base()
        {
            if (hasChildren)
            {
                var placeholder = new TreeViewItem() { Header = "{Please wait}" };
                this.Items.Add(placeholder);

                // Setup the expanded event
                this.Expanded += LoadSubtree;
            }
        }

        #endregion

        #region Public Instance Methods

        /// <summary>
        /// Creates source code text provided by the <paramref name="provider">CodeDomProvider</paramref>
        /// object that represents the selected object.
        /// </summary>
        /// <param name="provider">
        /// The <see cref="CodeDomProvider"/> to use when generating the source code.
        /// </param>
        /// <returns>
        /// The source code text provided by the <paramref name="provider">CodeDomProvider</paramref> 
        /// object that represents the selected object.
        /// </returns>
        public abstract string BuildCodeDomTextDocument(CodeDomProvider provider);

        /// <summary>
        /// Creates the xml source code of the currently selected object.
        /// </summary>
        /// <returns>
        /// The xml source code of the currently selected object.
        /// </returns>
        public virtual string BuildXmlTextDocument()
        {
            return String.Empty;
        }

        #endregion

        #region Protected Instance Methods

        /// <summary>
        /// Creates a new <see cref="OpenXmlElementTreeViewItem"/> collection based on the
        /// elements from <paramref name="elements"/>.
        /// </summary>
        /// <param name="elements">
        /// The existing collection of <see cref="OpenXmlElement"/> items.
        /// </param>
        /// <returns>
        /// A new collection of <see cref="OpenXmlElementTreeViewItem"/> elements.
        /// </returns>
        protected IEnumerable<OpenXmlElementTreeViewItem> BuildOpenXmlElementTreeViewItems(
            params OpenXmlElement[] elements)
        {
            return BuildOpenXmlElementTreeViewItems(elements.ToList());
        }

        /// <summary>
        /// Creates a new <see cref="OpenXmlElementTreeViewItem"/> collection based on the
        /// elements from <paramref name="elements"/>.
        /// </summary>
        /// <param name="elements">
        /// The existing collection of <see cref="OpenXmlElement"/> items.
        /// </param>
        /// <returns>
        /// A new collection of <see cref="OpenXmlElementTreeViewItem"/> elements.
        /// </returns>
        protected virtual IEnumerable<OpenXmlElementTreeViewItem> BuildOpenXmlElementTreeViewItems(
            IEnumerable<OpenXmlElement> elements)
        {
            if (elements is null || elements.Count() == 0)
            {
                throw new ArgumentNullException(nameof(elements));
            }
            string header;
            Row row;
            Cell cell;
            uint index = 0;

            foreach (var e in elements)
            {
                header = $"<{index++}> {e.LocalName} ({e.GetType().Name})";
                row = e as Row;
                cell = e as Cell;

                if (row != null && row.RowIndex != null && row.RowIndex.HasValue)
                {
                    header += $" [{(e as Row).RowIndex.Value}]";
                }
                else if (cell != null && cell.CellReference != null && cell.CellReference.HasValue)
                {
                    header += $" [{(e as Cell).CellReference.Value}]";
                }
                yield return new OpenXmlElementTreeViewItem(e)
                {
                    Header = header
                };
            }
        }

        /// <summary>
        /// Creates a new <see cref="OpenXmlPartTreeViewItem"/> collection based on the 
        /// parts from <paramref name="parts"/>.
        /// </summary>
        /// <param name="parts">
        /// The existing collection of <see cref="OpenXmlPart"/> items.
        /// </param>
        /// <returns>
        /// A new collection of <see cref="OpenXmlPartTreeViewItem"/> elements.
        /// </returns>
        protected IEnumerable<OpenXmlPartTreeViewItem> BuildOpenXmlPartTreeViewItems(
            params IdPartPair[] parts)
        {
            return BuildOpenXmlPartTreeViewItems(parts.ToList());
        }

        /// <summary>
        /// Creates a new <see cref="OpenXmlPartTreeViewItem"/> collection based on the 
        /// parts from <paramref name="parts"/>.
        /// </summary>
        /// <param name="parts">
        /// The existing collection of <see cref="OpenXmlPart"/> items.
        /// </param>
        /// <returns>
        /// A new collection of <see cref="OpenXmlPartTreeViewItem"/> elements.
        /// </returns>
        protected virtual IEnumerable<OpenXmlPartTreeViewItem> BuildOpenXmlPartTreeViewItems(
            IEnumerable<IdPartPair> parts)
        {
            foreach (var p in parts)
            {
                yield return new OpenXmlPartTreeViewItem(p)
                {
                    Header = $"[{p.RelationshipId}] {p.OpenXmlPart.Uri} ({p.OpenXmlPart.GetType().Name})"
                };
            }
        }

        /// <summary>
        /// Retrieves all of the <see cref="OpenXmlTreeViewItem"/> that should be directly
        /// added as child elements to the current object.
        /// </summary>
        /// <returns>
        /// A collection of <see cref="OpenXmlTreeViewItem"/> elements to add to the
        /// current object.
        /// </returns>
        protected abstract IEnumerable<OpenXmlTreeViewItem> GetOpenXmlSubtreeItems();

        #endregion

        #region Private Instance Methods

        /// <summary>
        /// Event triggered when the current tree item is expanded.
        /// </summary>
        /// <param name="sender">
        /// Object raising this event.
        /// </param>
        /// <param name="e">Event arguments</param>
        /// <remarks>
        /// This structure was setup to limit how much of the current openxml document
        /// is processed at a single time to reduce memory footprint.
        /// </remarks>
        private void LoadSubtree(object sender, EventArgs e)
        {
            if (_subtreeIsLoaded) return;

            var children = GetOpenXmlSubtreeItems();
            this.Items.Clear();

            foreach (var c in children)
            {
                Items.Add(c);
            }
            _subtreeIsLoaded = true;
        }

        #endregion
    }
}
