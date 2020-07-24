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

using ICSharpCode.AvalonEdit.Highlighting;
using Serialize.OpenXml.CodeGen;
using System;
using System.CodeDom.Compiler;

namespace DocxToSource.Wpf.Languages
{
    /// <summary>
    /// Base class for language definitions for the project.
    /// </summary>
    public abstract class LanguageDefinition
    {
        #region Protected Constructor

        /// <summary>
        /// Initializes a new instance of the <see cref="LanguageDefinition"/> class
        /// with the <see cref="CodeDomProvider"/> object that will be used to
        /// create the language source code.
        /// </summary>
        /// <param name="provider">
        /// The <see cref="CodeDomProvider"/> object for the new <see cref="LanguageDefinition"/>
        /// provider.
        /// </param>
        /// <exception cref="ArgumentNullException">
        /// <paramref name="provider"/> is <see langword="null"/>.
        /// </exception>
        protected LanguageDefinition(CodeDomProvider provider)
            : this(NamespaceAliasOptions.Default, provider) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="LanguageDefinition"/> class
        /// with the <see cref="CodeDomProvider"/> object that will be used to
        /// create the language source code and predefined <see cref="NamespaceAliasOptions"/>.
        /// </summary>
        /// <param name="opts">
        /// Custom <see cref="NamespaceAliasOptions"/> for the language definition.
        /// </param>
        /// <param name="provider">
        /// The <see cref="CodeDomProvider"/> object for the new <see cref="LanguageDefinition"/>
        /// provider.
        /// </param>
        /// <exception cref="ArgumentNullException">
        /// <paramref name="opts"/> or <paramref name="provider"/> is <see langword="null"/>.
        /// </exception>
        protected LanguageDefinition(NamespaceAliasOptions opts, CodeDomProvider provider)
        {
            Options = opts ?? throw new ArgumentNullException(nameof(opts));
            Provider = provider ?? throw new ArgumentNullException(nameof(provider));
            Highlighting = HighlightingManager.Instance.GetDefinitionByExtension("." + Provider.FileExtension);
        }

        #endregion

        #region Public Instance Properties

        /// <summary>
        /// Gets the name to display to the user.
        /// </summary>
        public string DisplayName { get; protected set; }

        /// <summary>
        /// Gets the <see cref="IHighlightingDefinition"/> to use for the generated
        /// source code.
        /// </summary>
        public IHighlightingDefinition Highlighting { get; private set; }

        /// <summary>
        /// Gets the <see cref="NamespaceAliasOptions"/> to use when generating
        /// source code.
        /// </summary>
        public NamespaceAliasOptions Options { get; private set; }

        /// <summary>
        /// Gets the <see cref="CodeDomProvider"/> to use when generating
        /// source code.
        /// </summary>
        public CodeDomProvider Provider { get; private set; }

        #endregion

        #region Public Instance Methods

        /// <inheritdoc/>
        public override string ToString() => DisplayName;

        #endregion
    }
}
