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
using DocxToSource.Wpf.Controls;
using DocxToSource.Wpf.Languages;
using ICSharpCode.AvalonEdit.Document;
using ICSharpCode.AvalonEdit.Highlighting;
using Microsoft.Win32;
using Prism.Commands;
using Prism.Mvvm;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.IO.Packaging;
using System.Windows;
using System.Windows.Input;

namespace DocxToSource.Wpf
{
    /// <summary>
    /// View model class for the main window.
    /// </summary>
    public class MainWindowModel : BindableBase, IDisposable
    {
        #region Private Static Fields

        /// <summary>
        /// Holds the default <see cref="IHighlightingDefinition"/> to use for
        /// the Xml window.
        /// </summary>
        private static readonly IHighlightingDefinition _defaultXmlDefinition;

        #endregion

        #region Private Instance Fields

        /// <summary>
        /// Holds the source code of the current selected openxml object.
        /// </summary>
        private TextDocument _codeDocument;

        /// <summary>
        /// Holds the highlighting definition for the source code text editor.
        /// </summary>
        private IHighlightingDefinition _codeSyntax;

        /// <summary>
        /// Holds the <see cref="DirectoryInfo"/> object containing the full path
        /// to use as the initial path in any OpenFileDialog windows.
        /// </summary>
        private DirectoryInfo _currentFileDirectory;

        /// <summary>
        /// Holds the full path and filename of the file that is currently open.
        /// </summary>
        private string _fileName;

        /// <summary>
        /// Indicates whether or not to automatically generate source code when 
        /// selecting DOM nodes.
        /// </summary>
        private bool _generateSourceCode;

        /// <summary>
        /// Indicates whether or not to enable syntax highlighting in the source code
        /// windows.
        /// </summary>
        private bool _highlightSyntax;

        /// <summary>
        /// Indicates whether or not the selected item represents an
        /// <see cref="OpenXmlElement"/> object.
        /// </summary>
        private bool _isOpenXmlElement;

        /// <summary>
        /// Holds the openxml file package the is currently being reviewed.
        /// </summary>
        private OpenXmlPackage _oPkg;

        /// <summary>
        /// Holds the raw package used to stage the stream information for 
        /// validation purposes.
        /// </summary>
        private Package _pkg;

        /// <summary>
        /// Holds the current treeviewitem that is currently selected in the treeview.
        /// </summary>
        private object _selectedItem;

        /// <summary>
        /// Holds the currently selected <see cref="LanguageDefinition"/> object.
        /// </summary>
        private LanguageDefinition _selectedLanguage;

        /// <summary>
        /// Holds the io stream containing the contents of the openxml file package.
        /// </summary>
        private Stream _stream;

        /// <summary>
        /// Holds the detailed exception information to display in the tree list view.
        /// </summary>
        private ObservableCollection<OpenXmlTreeViewItem> _treeData;

        /// <summary>
        /// Indicates whether or not to have the text in the source code windows word wrap.
        /// </summary>
        private bool _wordWrap;

        /// <summary>
        /// Holds the xml code fo the current selected openxml element.
        /// </summary>
        private TextDocument _xmlDocument;

        /// <summary>
        /// Holds the highlighting definition for the XML Text Editor
        /// </summary>
        private IHighlightingDefinition _xmlDocumentSyntax;

        #endregion

        #region Static Constructors

        /// <summary>
        /// Static Constructor.
        /// </summary>
        static MainWindowModel()
        {
            // Setup the default Xml definition to use when syntax highlighting is requested
            _defaultXmlDefinition = HighlightingManager.Instance.GetDefinition("XML");
        }

        #endregion

        #region Public Constructors

        /// <summary>
        /// Initializes a new instance of the <see cref="MainWindowModel"/> class that
        /// is empty.
        /// </summary>
        public MainWindowModel() : base()
        {
            _currentFileDirectory = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            _treeData = new ObservableCollection<OpenXmlTreeViewItem>();

            var langeDefs = new ObservableCollection<LanguageDefinition>();
            LanguageDefinitions = new ReadOnlyObservableCollection<LanguageDefinition>(langeDefs);

            // Load the language definition list
            langeDefs.Add(new CSharpLanguageDefinition());
            langeDefs.Add(new VBLanguageDefinition());
            //langeDefs.Add(new BooLanguageDefinition());

            // Set the default language
            SelectedLanguage = LanguageDefinitions[0];

            _codeDocument = new TextDocument();
            _xmlDocument = new TextDocument();

            CloseCommand = new DelegateCommand(() =>
            {
                Dispose();
                _treeData.Clear();
                RefreshSourceCodeWindows(null);
            });
            OpenCommand = new DelegateCommand(OpenOfficeDocument);

            QuitCommand = new DelegateCommand(() =>
            {
                Dispose();
                Application.Current.Shutdown();
            });
        }

        #endregion

        #region Public Instance Properties

        /// <summary>
        /// Gets the command to close the current document.
        /// </summary>
        public ICommand CloseCommand { get; private set; }

        /// <summary>
        /// Gets or sets the source code document object to display to the user.
        /// </summary>
        public TextDocument CodeDocument
        {
            get { return _codeDocument; }
            set
            {
                _codeDocument = value;
                FireChangeEvent(nameof(CodeDocument));
            }
        }

        /// <summary>
        /// Gets or sets the syntax highlighting definition for the source code text editor.
        /// </summary>
        public IHighlightingDefinition CodeDocumentSyntax
        {
            get { return _codeSyntax; }
            set
            {
                _codeSyntax = value;
                FireChangeEvent(nameof(CodeDocumentSyntax));
            }
        }

        /// <summary>
        /// Gets or sets the source code text to display to the user.
        /// </summary>
        public string CodeDocumentText
        {
            get { return _codeDocument.Text; }
            private set
            {
                _codeDocument.Text = value;
                FireChangeEvent(nameof(CodeDocument));
            }
        }

        /// <summary>
        /// Indicates whether or not to automatically generate source code when 
        /// selecting DOM nodes.
        /// </summary>
        public bool GenerateSourceCode
        {
            get { return _generateSourceCode; }
            set
            {
                _generateSourceCode = value;
                var item = GenerateSourceCode ? SelectedItem as OpenXmlTreeViewItem : null;
                RefreshSourceCodeWindows(item);
                FireChangeEvent(nameof(GenerateSourceCode));
            }
        }

        /// <summary>
        /// Indicates whether or not to enable syntax highlighting in the source code
        /// windows.
        /// </summary>
        public bool HighlightSyntax
        {
            get { return _highlightSyntax; }
            set
            {
                _highlightSyntax = value;
                ToggleSyntaxHighlighting(GenerateSourceCode && !(SelectedItem is null) && _highlightSyntax);
                FireChangeEvent(nameof(HighlightSyntax));
            }
        }

        /// <summary>
        /// Indicates whether or not the selected item represents an
        /// <see cref="OpenXmlElement"/> object.
        /// </summary>
        public bool IsOpenXmlElement
        {
            get { return _isOpenXmlElement; }
            set
            {
                _isOpenXmlElement = value;
                FireChangeEvent(nameof(IsOpenXmlElement));
            }
        }

        /// <summary>
        /// Gets the collection of <see cref="LanguageDefinition"/> objects that
        /// the user can select.
        /// </summary>
        public ReadOnlyObservableCollection<LanguageDefinition> LanguageDefinitions
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the command to open a new office 2007+ document.
        /// </summary>
        public ICommand OpenCommand { get; private set; }

        /// <summary>
        /// Gets the command that shuts down the application.
        /// </summary>
        public ICommand QuitCommand { get; private set; }

        /// <summary>
        /// Gets or sets the <see cref="OpenXmlTreeViewItem"/> that is currently selected.
        /// </summary>
        public object SelectedItem
        {
            get { return _selectedItem; }
            set
            {
                _selectedItem = value;

                if (GenerateSourceCode && _selectedItem is OpenXmlTreeViewItem)
                {
                    RefreshSourceCodeWindows(_selectedItem as OpenXmlTreeViewItem);
                }
                FireChangeEvent(nameof(SelectedItem));
            }
        }

        /// <summary>
        /// Gets or sets the <see cref="LanguageDefinition"/> object currently
        /// selected by the user.
        /// </summary>
        public LanguageDefinition SelectedLanguage
        {
            get { return _selectedLanguage; }
            set
            {
                _selectedLanguage = value;
                if (GenerateSourceCode)
                {
                    RefreshSourceCodeWindows(SelectedItem as OpenXmlTreeViewItem);
                }
                FireChangeEvent(nameof(SelectedLanguage));
            }
        }

        /// <summary>
        /// Gets all of the openxml objects to display in the tree.
        /// </summary>
        public ObservableCollection<OpenXmlTreeViewItem> TreeData
        {
            get { return _treeData; }
            private set
            {
                _treeData = value;
                FireChangeEvent(nameof(TreeData));
            }
        }

        /// <summary>
        /// Indicates whether or not to have the text in the source code windows word wrap.
        /// </summary>
        public bool WordWrap
        {
            get { return _wordWrap; }
            set
            {
                _wordWrap = value;
                FireChangeEvent(nameof(WordWrap));
            }
        }

        /// <summary>
        /// Gets or sets the source code to display to the user.
        /// </summary>
        public TextDocument XmlSourceDocument
        {
            get { return _xmlDocument; }
            set
            {
                _xmlDocument = value;
                FireChangeEvent(nameof(XmlSourceDocument));
            }
        }

        /// <summary>
        /// Gets or sets the syntax highlighting definition for the xml document text editor.
        /// </summary>
        public IHighlightingDefinition XmlSourceDocumentSyntax
        {
            get { return _xmlDocumentSyntax; }
            set
            {
                _xmlDocumentSyntax = value;
                FireChangeEvent(nameof(XmlSourceDocumentSyntax));
            }
        }

        /// <summary>
        /// Gets or sets the source code text to display to the user.
        /// </summary>
        public string XmlSourceDocumentText
        {
            get { return _xmlDocument.Text; }
            set
            {
                _xmlDocument.Text = value;
                FireChangeEvent(nameof(XmlSourceDocument));
            }
        }

        #endregion

        #region Public Instance Methods

        /// <summary>
        /// Method to make sure that all unmanaged resources are released properly.
        /// </summary>
        public void Dispose()
        {
            if (_oPkg != null)
            {
                _oPkg.Close();
                _oPkg.Dispose();
                _oPkg = null;
            }
            if (_pkg != null)
            {
                _pkg.Close();
                _pkg = null;
            }
            if (_stream != null)
            {
                _stream.Close();
                _stream.Dispose();
                _stream = null;
            }
        }

        #endregion

        #region Private Instance Methods

        /// <summary>
        /// Shortcut method to raise the <see cref="PropertyChangedEventHandler"/>
        /// event.
        /// </summary>
        /// <param name="name">
        /// Name of the property raising the event.
        /// </param>
        private void FireChangeEvent(string name) =>
            OnPropertyChanged(new PropertyChangedEventArgs(name));

        /// <summary>
        /// Resets the main window controls and loads a requested OpenXml based file.
        /// </summary>
        private void OpenOfficeDocument()
        {
            const string docxIdUri = "/word/document.xml";
            const string xlsxIdUri = "/xl/workbook.xml";
            const string pptxIdUri = "/ppt/presentation.xml";
            const string fileFilter =
                "All Microsoft Office 2007+ valid documents (*.xlsx;*.xlsm;*.pptx;*.pptm;*.docx;*.docm)|*.xlsx;*.xlsm;*.pptx;*.pptm;*.docx;*.docm" +
                "|Microsoft Excel 2007+ documents (*.xlsx;*.xlsm)|*.xlsx;*.xlsm" +
                "|Microsoft Powerpoint 2007+ documents (*.pptx;*.pptm)|*.pptx;*.pptm" +
                "|Microsoft Word 2007+ documents (*.docx;*.docm)|*.docx;*.docm" +
                "|All files | *.*";

            bool? dialogResult;
            var ofDialog = new OpenFileDialog()
            {
                InitialDirectory = _currentFileDirectory.FullName,
                Multiselect = false,
                CheckFileExists = true,
                CheckPathExists = true,
                ReadOnlyChecked = true,
                Filter = fileFilter,
                FilterIndex = 1
            };

            dialogResult = ofDialog.ShowDialog();
            if (!dialogResult.GetValueOrDefault(false))
            {
                // If the user cancels out; exit method.
                return;
            }

            // Ensure that everything is cleared out before proceeding
            Dispose();
            CodeDocument.FileName = null;
            XmlSourceDocument.FileName = null;
            CodeDocumentText = String.Empty;
            XmlSourceDocumentText = String.Empty;
            _fileName = String.Empty;

            // See if the user left the Open as readonly flag on
            var fShare = ofDialog.ReadOnlyChecked ? FileShare.ReadWrite : FileShare.Read;

            // Get the selected file details
            var fi = new FileInfo(ofDialog.FileName);
            _currentFileDirectory = fi.Directory;
            _fileName = fi.Name;
            _stream = fi.Open(FileMode.Open, FileAccess.Read, fShare);
            _pkg = Package.Open(_stream);

            // Setup a quick look up for easier package validation
            var quickPicks = new Dictionary<string, Func<Package, OpenXmlPackage>>(3)
            {
                { docxIdUri, WordprocessingDocument.Open },
                { xlsxIdUri, SpreadsheetDocument.Open },
                { pptxIdUri, PresentationDocument.Open }
            };

            foreach (var qp in quickPicks)
            {
                if (_pkg.PartExists(new Uri(qp.Key, UriKind.Relative)))
                {
                    _oPkg = qp.Value.Invoke(_pkg);
                    break;
                }
            }

            // Make sure that a valid package was found before proceeding.
            if (_oPkg == null)
            {
                throw new InvalidDataException("Selected file is not a known/valid OpenXml document");
            }

            // Wrap it up
            var mainItem = new OpenXmlPackageTreeViewItem(_oPkg) { Header = _fileName };
            _treeData.Clear();
            _treeData.Add(mainItem);
        }

        /// <summary>
        /// Refreshes the <see cref="ICSharpCode.AvalonEdit.TextEditor"/> controls
        /// in the main window.
        /// </summary>
        /// <param name="item">
        /// The <see cref="OpenXmlTreeViewItem"/> currently selected by the user.
        /// </param>
        /// <remarks>
        /// Passing <see langword="null"/> as the <paramref name="item"/> will cause
        /// the <see cref="ICSharpCode.AvalonEdit.TextEditor"/> controls to clear their
        /// contents.
        /// </remarks>
        private void RefreshSourceCodeWindows(OpenXmlTreeViewItem item)
        {
            if (item is null)
            {
                CodeDocument.FileName = null;
                XmlSourceDocument.FileName = null;
                CodeDocumentText = String.Empty;
                XmlSourceDocumentText = String.Empty;
                ToggleSyntaxHighlighting(false);
            }
            else
            {
                var randName = Path.GetFileNameWithoutExtension(Path.GetRandomFileName());
                CodeDocument.FileName = randName + "." + SelectedLanguage.Provider.FileExtension;
                XmlSourceDocument.FileName = randName + ".xml";

                CodeDocumentText = item.BuildCodeDomTextDocument(SelectedLanguage.Provider);
                XmlSourceDocumentText = item.BuildXmlTextDocument();

                ToggleSyntaxHighlighting(HighlightSyntax);
            }
            IsOpenXmlElement = !String.IsNullOrWhiteSpace(XmlSourceDocumentText);
            FireChangeEvent(nameof(IsOpenXmlElement));
        }

        /// <summary>
        /// Enables or disables syntax highlighting in the code windows.
        /// </summary>
        /// <param name="enable">
        /// <see langword="true"/> to turn on syntax highlighting; <see langword="false"/>
        /// to turn it off.
        /// </param>
        private void ToggleSyntaxHighlighting(bool enable)
        {
            CodeDocumentSyntax = enable ? SelectedLanguage.Highlighting : null;
            XmlSourceDocumentSyntax = enable ? _defaultXmlDefinition : null;
        }

        #endregion
    }
}
