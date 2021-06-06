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
using System.CodeDom;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;

namespace Serialize.OpenXml.CodeGen
{
    /// <summary>
    /// Defines objects that provide custom code generation instructions for
    /// <see cref="OpenXmlPart"/> derived objects that the process may encounter.
    /// </summary>
    public interface IOpenXmlPartHandler : IOpenXmlHandler
    {
        #region Methods

        /// <summary>
        /// Creates the appropriate code objects needed to create the entry method for the
        /// current request.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object and relationship id to build code for.
        /// </param>
        /// <param name="settings">
        /// The <see cref="ISerializeSettings"/> to use during the code generation
        /// process.
        /// </param>
        /// /// <param name="typeCounts">
        /// A lookup <see cref="IDictionary{TKey, TValue}"/> object containing the
        /// number of times a given type was referenced.  This is used for variable naming
        /// purposes.
        /// </param>
        /// <param name="namespaces">
        /// Collection <see cref="ISet{T}"/> used to keep track of all openxml namespaces
        /// used during the process.
        /// </param>
        /// <param name="blueprints">
        /// The collection of <see cref="OpenXmlPartBluePrint"/> objects that have already been
        /// visited.
        /// </param>
        /// <param name="rootVar">
        /// The root variable name and <see cref="Type"/> to use when building code
        /// statements to create new <see cref="OpenXmlPart"/> objects.
        /// </param>
        /// <returns>
        /// A collection of code statements and expressions that could be used to generate
        /// a new <paramref name="part"/> object from code.
        /// </returns>
        /// <remarks>
        /// If this method returns <see langword="null"/>, the default implementation will
        /// be used instead.
        /// </remarks>
        CodeStatementCollection BuildEntryMethodCodeStatements(
            IdPartPair part,
            ISerializeSettings settings,
            IDictionary<string, int> typeCounts,
            ISet<string> namespaces,
            OpenXmlPartBluePrintCollection blueprints,
            KeyValuePair<string, Type> rootVar);

        /// <summary>
        /// Builds the helper method of a given <see cref="OpenXmlPart"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate the source code for.
        /// </param>
        /// <param name="methodName">
        /// The name of the method to use when building the 
        /// </param>
        /// <param name="settings">
        /// The <see cref="ISerializeSettings"/> to use during the code generation
        /// process.
        /// </param>
        /// <param name="namespaces">
        /// Collection <see cref="ISet{T}"/> used to keep track of all openxml namespaces
        /// used during the process.
        /// </param>
        /// <returns>
        /// A new <see cref="CodeMemberMethod"/> object containing the necessary source code
        /// to recreate <paramref name="part"/>.
        /// </returns>
        /// <remarks>
        /// If this method returns <see langword="null"/>, the default implementation will
        /// be used instead.
        /// </remarks>
        CodeMemberMethod BuildHelperMethod(
            OpenXmlPart part,
            string methodName,
            ISerializeSettings settings,
            ISet<string> namespaces);

        #endregion
    }
}
