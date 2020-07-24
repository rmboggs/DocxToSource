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

using System.Windows;
using System.Windows.Controls;

namespace DocxToSource.Wpf.Controls
{
    /// <summary>
    /// A subclass of the default wpf <see cref="System.Windows.Controls.TreeView"/> designed
    /// specifically for objects from the DocumentFormat.OpenXml library.
    /// </summary>
    public class OpenXmlTreeView : TreeView
    {
        #region Public Static Fields

        /// <summary>
        /// Static dependency property handler for the SelectedTreeItem property.
        /// </summary>
        public static readonly DependencyProperty SelectedItem_Property =
            DependencyProperty.Register("SelectedTreeItem", typeof(object),
                typeof(OpenXmlTreeView), new UIPropertyMetadata(null));

        #endregion

        #region Public Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="OpenXmlTreeView"/> class that
        /// is empty.
        /// </summary>
        public OpenXmlTreeView() : base()
        {
            this.SelectedItemChanged +=
                new RoutedPropertyChangedEventHandler<object>(MapSelecteditem);
        }

        #endregion

        #region Public Instance Properties

        /// <summary>
        /// Gets or sets the item currently selected in the tree.
        /// </summary>
        public object SelectedTreeItem
        {
            get { return (object)GetValue(SelectedItem_Property); }
            set { SetValue(SelectedItem_Property, value); }
        }

        #endregion

        #region Private Instance Methods

        /// <summary>
        /// Maps the selected item object to the selected tree item property.
        /// </summary>
        /// <param name="sender">
        /// Object sending event.
        /// </param>
        /// <param name="e">
        /// Event args
        /// </param>
        private void MapSelecteditem(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            if (SelectedItem != null)
            {
                SetValue(SelectedItem_Property, SelectedItem);
            }
        }

        #endregion
    }
}
