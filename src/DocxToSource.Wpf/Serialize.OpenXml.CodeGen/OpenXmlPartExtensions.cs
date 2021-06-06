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
using Serialize.OpenXml.CodeGen.Extentions;
using System;
using System.CodeDom;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Serialize.OpenXml.CodeGen
{
    /// <summary>
    /// Static class that converts <see cref="OpenXmlPart">OpenXmlParts</see>
    /// into Code DOM objects.
    /// </summary>
    public static class OpenXmlPartExtensions
    {
        #region Private Static Fields

        /// <summary>
        /// The default parameter name for an <see cref="OpenXmlPart"/> object.
        /// </summary>
        private const string methodParamName = "part";

        #endregion

        #region Public Static Methods

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
        public static CodeStatementCollection BuildEntryMethodCodeStatements(
            IdPartPair part,
            ISerializeSettings settings,
            IDictionary<string, int> typeCounts,
            ISet<string> namespaces,
            OpenXmlPartBluePrintCollection blueprints,
            KeyValuePair<string, Type> rootVar)
        {
            // Argument validation
            if (part is null) throw new ArgumentNullException(nameof(part));
            if (settings is null) throw new ArgumentNullException(nameof(settings));
            if (blueprints is null) throw new ArgumentNullException(nameof(blueprints));
            if (String.IsNullOrWhiteSpace(rootVar.Key)) throw new ArgumentNullException(nameof(rootVar.Key));
            bool hasHandlers = settings?.Handlers != null;

            // Use the custom handler methods if present and provide actual code
            if (hasHandlers && settings.Handlers.TryGetValue(part.OpenXmlPart.GetType(), out IOpenXmlHandler h))
            {
                if (h is IOpenXmlPartHandler partHandler)
                {
                    var customStatements = partHandler.BuildEntryMethodCodeStatements(
                        part, settings, typeCounts, namespaces, blueprints, rootVar);

                    if (customStatements != null) return customStatements;
                }
            }

            var result = new CodeStatementCollection();
            var partType = part.OpenXmlPart.GetType();
            var partTypeName = partType.Name;
            var partTypeFullName = partType.FullName;
            string varName = partType.Name.ToCamelCase();
            bool customAddNewPartRequired = CheckForCustomAddNewPartMethod(partType, rootVar.Value, out string addNewPartName);

#pragma warning disable IDE0018 // Inline variable declaration
            OpenXmlPartBluePrint bpTemp;
#pragma warning restore IDE0018 // Inline variable declaration

            CodeMethodReferenceExpression referenceExpression;
            CodeMethodInvokeExpression invokeExpression;
            CodeMethodReferenceExpression methodReference;

            // Make sure that the namespace for the current part is captured
            namespaces.Add(partType.Namespace);

            // If the URI of the current part has already been included into
            // the blue prints collection, build the AddPart invocation
            // code statement and exit current method iteration.
            if (blueprints.TryGetValue(part.OpenXmlPart.Uri, out bpTemp))
            {
                // Surround this snippet with blank lines to make it
                // stand out in the current section of code.
                result.AddBlankLine();
                referenceExpression = new CodeMethodReferenceExpression(
                    new CodeVariableReferenceExpression(rootVar.Key), "AddPart",
                    new CodeTypeReference(part.OpenXmlPart.GetType().Name));
                invokeExpression = new CodeMethodInvokeExpression(referenceExpression,
                    new CodeVariableReferenceExpression(bpTemp.VariableName),
                    new CodePrimitiveExpression(part.RelationshipId));
                result.Add(invokeExpression);
                result.AddBlankLine();
                return result;
            }

            // Assign the appropriate variable name
            if (typeCounts.ContainsKey(partTypeFullName))
            {
                varName = String.Concat(varName, typeCounts[partTypeFullName]++);
            }
            else
            {
                typeCounts.Add(partTypeFullName, 1);
            }

            // Setup the blueprint
            bpTemp = new OpenXmlPartBluePrint(part.OpenXmlPart, varName);

            // Need to evaluate the current OpenXmlPart type first to make sure the
            // correct "Add" statement is used as not all Parts can be initialized
            // using the "AddNewPart" method

            // Check for image part methods
            if (customAddNewPartRequired)
            {
                referenceExpression = new CodeMethodReferenceExpression(
                    new CodeVariableReferenceExpression(rootVar.Key), addNewPartName);
            }
            else
            {
                // Setup the add new part statement for the current OpenXmlPart object
                referenceExpression = new CodeMethodReferenceExpression(
                    new CodeVariableReferenceExpression(rootVar.Key), "AddNewPart",
                    new CodeTypeReference(partTypeName));
            }

            // Create the invoke expression
            invokeExpression = new CodeMethodInvokeExpression(referenceExpression);

            // Add content type to invoke method for embeddedpackage and image parts.
            if (part.OpenXmlPart is EmbeddedPackagePart || customAddNewPartRequired)
            {
                invokeExpression.Parameters.Add(
                    new CodePrimitiveExpression(part.OpenXmlPart.ContentType));
            }
            else if (!customAddNewPartRequired)
            {
                invokeExpression.Parameters.Add(
                    new CodePrimitiveExpression(part.RelationshipId));
            }

            result.Add(new CodeVariableDeclarationStatement(partTypeName, varName, invokeExpression));

            // Because the custom AddNewPart methods don't consistently take in a string relId
            // as a parameter, the id needs to be assigned after it is created.
            if (customAddNewPartRequired)
            {
                methodReference = new CodeMethodReferenceExpression(
                    new CodeVariableReferenceExpression(rootVar.Key), "ChangeIdOfPart");
                result.Add(new CodeMethodInvokeExpression(methodReference,
                    new CodeVariableReferenceExpression(varName),
                    new CodePrimitiveExpression(part.RelationshipId)));
            }

            // Add the call to the method to populate the current OpenXmlPart object
            methodReference = new CodeMethodReferenceExpression(new CodeThisReferenceExpression(), bpTemp.MethodName);
            result.Add(new CodeMethodInvokeExpression(methodReference,
                new CodeDirectionExpression(FieldDirection.Ref,
                    new CodeVariableReferenceExpression(varName))));

            // Add the appropriate code statements if the current part
            // contains any hyperlink relationships
            if (part.OpenXmlPart.HyperlinkRelationships.Count() > 0)
            {
                // Add a line break first for easier reading
                result.AddBlankLine();
                result.AddRange(
                    part.OpenXmlPart.HyperlinkRelationships.BuildHyperlinkRelationshipStatements(varName));
            }

            // Add the appropriate code statements if the current part
            // contains any non-hyperlink external relationships
            if (part.OpenXmlPart.ExternalRelationships.Count() > 0)
            {
                // Add a line break first for easier reading
                result.AddBlankLine();
                result.AddRange(
                    part.OpenXmlPart.ExternalRelationships.BuildExternalRelationshipStatements(varName));
            }

            // put a line break before going through the child parts
            result.AddBlankLine();

            // Add the current blueprint to the collection
            blueprints.Add(bpTemp);

            // Now check to see if there are any child parts for the current OpenXmlPart object.
            if (bpTemp.Part.Parts != null)
            {
                OpenXmlPartBluePrint childBluePrint;

                foreach (var p in bpTemp.Part.Parts)
                {
                    // If the current child object has already been created, simply add a reference to
                    // said object using the AddPart method.
                    if (blueprints.Contains(p.OpenXmlPart.Uri))
                    {
                        childBluePrint = blueprints[p.OpenXmlPart.Uri];

                        referenceExpression = new CodeMethodReferenceExpression(
                            new CodeVariableReferenceExpression(varName), "AddPart",
                            new CodeTypeReference(p.OpenXmlPart.GetType().Name));

                        invokeExpression = new CodeMethodInvokeExpression(referenceExpression,
                            new CodeVariableReferenceExpression(childBluePrint.VariableName),
                            new CodePrimitiveExpression(p.RelationshipId));

                        result.Add(invokeExpression);
                        continue;
                    }

                    // If this is a new part, call this method with the current part's details
                    result.AddRange(BuildEntryMethodCodeStatements(p, settings, typeCounts, namespaces, blueprints,
                        new KeyValuePair<string, Type>(varName, partType)));
                }
            }

            return result;
        }

        /// <summary>
        /// Creates the appropriate helper methods for all of the <see cref="OpenXmlPart"/> objects
        /// for the current request.
        /// </summary>
        /// <param name="bluePrints">
        /// The collection of <see cref="OpenXmlPartBluePrint"/> objects that have already been
        /// visited.
        /// </param>
        /// <param name="settings">
        /// The <see cref="ISerializeSettings"/> to use during the code generation
        /// process.
        /// </param>
        /// <param name="namespaces">
        /// <see cref="ISet{T}"/> collection used to keep track of all openxml namespaces
        /// used during the process.
        /// </param>
        /// <returns>
        /// A collection of code helper statements and expressions that could be used to generate a new
        /// <see cref="OpenXmlPart"/> object from code.
        /// </returns>
        public static CodeTypeMemberCollection BuildHelperMethods(
            OpenXmlPartBluePrintCollection bluePrints,
            ISerializeSettings settings,
            ISet<string> namespaces)
        {
            if (bluePrints == null) throw new ArgumentNullException(nameof(bluePrints));
            var result = new CodeTypeMemberCollection();
            var localTypeCounts = new Dictionary<Type, int>();
            Type rootElementType;
            CodeMemberMethod method;
            bool hasHandlers = settings?.Handlers != null;

            foreach (var bp in bluePrints)
            {
                // Implement the custom helper if present
                if (hasHandlers && settings.Handlers.TryGetValue(bp.PartType, out IOpenXmlHandler h))
                {
                    if (h is IOpenXmlPartHandler partHandler)
                    {
                        method = partHandler.BuildHelperMethod(
                            bp.Part, bp.MethodName, settings, namespaces);

                        if (method != null)
                        {
                            result.Add(method);
                            continue;
                        }
                    }
                }

                // Setup the first method
                method = new CodeMemberMethod()
                {
                    Name = bp.MethodName,
                    Attributes = MemberAttributes.Private | MemberAttributes.Final,
                    ReturnType = new CodeTypeReference()
                };
                method.Parameters.Add(
                    new CodeParameterDeclarationExpression(bp.Part.GetType().Name, methodParamName)
                    { Direction = FieldDirection.Ref });

                // Code part elements next
                if (bp.Part.RootElement is null)
                {
                    // Put all of the pieces together
                    method.Statements.AddRange(bp.Part.BuildPartFeedData(namespaces));
                }
                else
                {
                    rootElementType = bp.Part.RootElement?.GetType();

                    // Build the element details of the requested part for the current method
                    method.Statements.AddRange(
                        bp.Part.RootElement.BuildCodeStatements(
                            settings, localTypeCounts, namespaces, out string rootElementVar));

                    // Now finish up the current method by assigning the OpenXmlElement code statements
                    // back to the appropriate property of the part parameter
                    if (rootElementType != null && !String.IsNullOrWhiteSpace(rootElementVar))
                    {
                        foreach (var paramProp in bp.Part.GetType().GetProperties())
                        {
                            if (paramProp.PropertyType == rootElementType)
                            {
                                var varRef = new CodeVariableReferenceExpression(rootElementVar);
                                var paramRef = new CodeArgumentReferenceExpression(methodParamName);
                                var propRef = new CodePropertyReferenceExpression(paramRef, paramProp.Name);
                                method.Statements.Add(new CodeAssignStatement(propRef, varRef));
                                break;
                            }
                        }
                    }
                }
                result.Add(method);
            }

            return result;
        }

        /// <summary>
        /// Creates a collection of code statements that describe how to add external relationships to
        /// a <see cref="OpenXmlPartContainer"/> object.
        /// </summary>
        /// <param name="relationships">
        /// The collection of <see cref="ExternalRelationship"/> objects to build the code statements for.
        /// </param>
        /// <param name="parentName">
        /// The name of the <see cref="OpenXmlPartContainer"/> object that the external relationship
        /// assignments should be for.
        /// </param>
        /// <returns>
        /// A collection of code statements that could be used to generate and assign new
        /// <see cref="ExternalRelationship"/> objects to a <see cref="OpenXmlPartContainer"/> object.
        /// </returns>
        public static CodeStatementCollection BuildExternalRelationshipStatements(
            this IEnumerable<ExternalRelationship> relationships, string parentName)
        {
            if (String.IsNullOrWhiteSpace(parentName)) throw new ArgumentNullException(nameof(parentName));

            var result = new CodeStatementCollection();

            // Return an empty code statement collection if the hyperlinks parameter is empty.
            if (relationships.Count() == 0) return result;

            CodeObjectCreateExpression createExpression;
            CodeMethodReferenceExpression methodReferenceExpression;
            CodeMethodInvokeExpression invokeExpression;
            CodePrimitiveExpression param;
            CodeTypeReference typeReference;

            foreach (var ex in relationships)
            {
                // Need special care to create the uri for the current object.
                typeReference = new CodeTypeReference(ex.Uri.GetType());
                param = new CodePrimitiveExpression(ex.Uri.ToString());
                createExpression = new CodeObjectCreateExpression(typeReference, param);

                // Create the AddHyperlinkRelationship statement
                methodReferenceExpression = new CodeMethodReferenceExpression(
                    new CodeVariableReferenceExpression(parentName),
                    "AddExternalRelationship");
                invokeExpression = new CodeMethodInvokeExpression(methodReferenceExpression,
                    new CodePrimitiveExpression(ex.RelationshipType),
                    createExpression,
                    new CodePrimitiveExpression(ex.Id));
                result.Add(invokeExpression);
            }
            return result;
        }

        /// <summary>
        /// Creates a collection of code statements that describe how to add hyperlink relationships to
        /// a <see cref="OpenXmlPartContainer"/> object.
        /// </summary>
        /// <param name="hyperlinks">
        /// The collection of <see cref="HyperlinkRelationship"/> objects to build the code statements for.
        /// </param>
        /// <param name="parentName">
        /// The name of the <see cref="OpenXmlPartContainer"/> object that the hyperlink relationship
        /// assignments should be for.
        /// </param>
        /// <returns>
        /// A collection of code statements that could be used to generate and assign new
        /// <see cref="HyperlinkRelationship"/> objects to a <see cref="OpenXmlPartContainer"/> object.
        /// </returns>
        public static CodeStatementCollection BuildHyperlinkRelationshipStatements(
            this IEnumerable<HyperlinkRelationship> hyperlinks, string parentName)
        {
            if (String.IsNullOrWhiteSpace(parentName)) throw new ArgumentNullException(nameof(parentName));

            var result = new CodeStatementCollection();

            // Return an empty code statement collection if the hyperlinks parameter is empty.
            if (hyperlinks.Count() == 0) return result;

            CodeObjectCreateExpression createExpression;
            CodeMethodReferenceExpression methodReferenceExpression;
            CodeMethodInvokeExpression invokeExpression;
            CodePrimitiveExpression param;
            CodeTypeReference typeReference;

            foreach (var hl in hyperlinks)
            {
                // Need special care to create the uri for the current object.
                typeReference = new CodeTypeReference(hl.Uri.GetType());
                param = new CodePrimitiveExpression(hl.Uri.ToString());
                createExpression = new CodeObjectCreateExpression(typeReference, param);

                // Create the AddHyperlinkRelationship statement
                methodReferenceExpression = new CodeMethodReferenceExpression(
                    new CodeVariableReferenceExpression(parentName),
                    "AddHyperlinkRelationship");
                invokeExpression = new CodeMethodInvokeExpression(methodReferenceExpression,
                    createExpression,
                    new CodePrimitiveExpression(hl.IsExternal),
                    new CodePrimitiveExpression(hl.Id));
                result.Add(invokeExpression);
            }
            return result;
        }

        /// <summary>
        /// Builds code statements that will build <paramref name="part"/> using the
        /// <see cref="OpenXmlPart.FeedData(Stream)"/> method.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to build the source code for.
        /// </param>
        /// <param name="namespaces">
        /// <see cref="ISet{T}">Set</see> of <see cref="String"/> values used to keep
        /// track of all openxml namespaces used during the process.
        /// </param>
        /// <returns>
        /// A <see cref="CodeStatementCollection">collection of code statements</see>
        /// that would regenerate <paramref name="part"/> using the
        /// <see cref="OpenXmlPart.FeedData(Stream)"/> method.
        /// </returns>
        public static CodeStatementCollection BuildPartFeedData(this OpenXmlPart part, ISet<string> namespaces)
        {
            // Make sure no null values were passed.
            if (part == null) throw new ArgumentNullException(nameof(part));
            if (namespaces == null) throw new ArgumentNullException(nameof(namespaces));

            // If the root element is not present (aka: null) then perform a simple feed
            // dump of the part in the current method
            const string memName = "mem";
            const string b64Name = "base64";

            var result = new CodeStatementCollection();

            // Add the necessary namespaces by hand to the namespace set
            namespaces.Add("System");
            namespaces.Add("System.IO");

            using (var partStream = part.GetStream(FileMode.Open, FileAccess.Read))
            {
                using (var mem = new MemoryStream())
                {
                    partStream.CopyTo(mem);
                    result.Add(new CodeVariableDeclarationStatement(typeof(string), b64Name,
                        new CodePrimitiveExpression(Convert.ToBase64String(mem.ToArray()))));
                }
            }
            result.AddBlankLine();

            var fromBase64 = new CodeMethodReferenceExpression(new CodeVariableReferenceExpression("Convert"),
                "FromBase64String");
            var invokeFromBase64 = new CodeMethodInvokeExpression(fromBase64, new CodeVariableReferenceExpression("base64"));
            var createStream = new CodeObjectCreateExpression(new CodeTypeReference("MemoryStream"),
                invokeFromBase64, new CodePrimitiveExpression(false));
            var feedData = new CodeMethodReferenceExpression(new CodeArgumentReferenceExpression(methodParamName), "FeedData");
            var invokeFeedData = new CodeMethodInvokeExpression(feedData, new CodeVariableReferenceExpression(memName));
            var disposeMem = new CodeMethodReferenceExpression(new CodeVariableReferenceExpression(memName), "Dispose");
            var invokeDisposeMem = new CodeMethodInvokeExpression(disposeMem);

            // Setup the try statement
            var tryAndCatch = new CodeTryCatchFinallyStatement();
            tryAndCatch.TryStatements.Add(invokeFeedData);
            tryAndCatch.FinallyStatements.Add(invokeDisposeMem);

            // Put all of the pieces together
            result.Add(new CodeVariableDeclarationStatement("Stream", memName, createStream));
            result.Add(tryAndCatch);

            return result;
        }

        /// <summary>
        /// Converts an <see cref="OpenXmlPart"/> into a CodeDom object that can be used
        /// to build code in a given .NET language to build the referenced <see cref="OpenXmlPart"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate source code for.
        /// </param>
        /// <returns>
        /// A new <see cref="CodeCompileUnit"/> containing the instructions to build
        /// the referenced <see cref="OpenXmlPart"/>.
        /// </returns>
        public static CodeCompileUnit GenerateSourceCode(this OpenXmlPart part)
        {
            return part.GenerateSourceCode(new DefaultSerializeSettings());
        }

        /// <summary>
        /// Converts an <see cref="OpenXmlPart"/> into a CodeDom object that can be used
        /// to build code in a given .NET language to build the referenced <see cref="OpenXmlPart"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate source code for.
        /// </param>
        /// <param name="opts">
        /// The <see cref="NamespaceAliasOptions"/> to apply to the resulting source code.
        /// </param>
        /// <returns>
        /// A new <see cref="CodeCompileUnit"/> containing the instructions to build
        /// the referenced <see cref="OpenXmlPart"/>.
        /// </returns>
        public static CodeCompileUnit GenerateSourceCode(this OpenXmlPart part, NamespaceAliasOptions opts)
        {
            return part.GenerateSourceCode(new DefaultSerializeSettings(opts));
        }

        /// <summary>
        /// Converts an <see cref="OpenXmlPart"/> into a CodeDom object that can be used
        /// to build code in a given .NET language to build the referenced <see cref="OpenXmlPart"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate source code for.
        /// </param>
        /// <param name="settings">
        /// The <see cref="ISerializeSettings"/> to use during the code generation
        /// process.
        /// </param>
        /// <returns>
        /// A new <see cref="CodeCompileUnit"/> containing the instructions to build
        /// the referenced <see cref="OpenXmlPart"/>.
        /// </returns>
        public static CodeCompileUnit GenerateSourceCode(this OpenXmlPart part, ISerializeSettings settings)
        {
            CodeMethodReferenceExpression methodRef;
            OpenXmlPartBluePrint mainBluePrint;
            var result = new CodeCompileUnit();
            var eType = part.GetType();
            var partTypeName = eType.Name;
            var partTypeFullName = eType.FullName;
            var varName = eType.Name.ToCamelCase();
            var partTypeCounts = new Dictionary<string, int>();
            var namespaces = new SortedSet<string>();
            var mainNamespace = new CodeNamespace("OpenXmlSample");
            var bluePrints = new OpenXmlPartBluePrintCollection();

            // Assign the appropriate variable name
            if (partTypeCounts.ContainsKey(partTypeFullName))
            {
                varName = String.Concat(varName, partTypeCounts[partTypeFullName]++);
            }
            else
            {
                partTypeCounts.Add(partTypeFullName, 1);
            }

            // Generate a new blue print for the current part to help create the main
            // method reference then add it to the blue print collection
            mainBluePrint = new OpenXmlPartBluePrint(part, varName);
            bluePrints.Add(mainBluePrint);
            methodRef = new CodeMethodReferenceExpression(new CodeThisReferenceExpression(), mainBluePrint.MethodName);

            // Build the entry method
            var entryMethod = new CodeMemberMethod()
            {
                Name = $"Create{partTypeName}",
                ReturnType = new CodeTypeReference(),
                Attributes = MemberAttributes.Public | MemberAttributes.Final
            };
            entryMethod.Parameters.Add(
                new CodeParameterDeclarationExpression(partTypeName, methodParamName)
                { Direction = FieldDirection.Ref });

            // Add all of the child part references here
            if (part.Parts != null)
            {
                var rootPartPair = new KeyValuePair<string, Type>(methodParamName, eType);
                foreach (var pair in part.Parts)
                {
                    entryMethod.Statements.AddRange(BuildEntryMethodCodeStatements(
                        pair, settings, partTypeCounts, namespaces, bluePrints, rootPartPair));
                }
            }

            entryMethod.Statements.Add(new CodeMethodInvokeExpression(methodRef,
                new CodeArgumentReferenceExpression(methodParamName)));

            // Setup the main class next
            var mainClass = new CodeTypeDeclaration($"{eType.Name}BuilderClass")
            {
                IsClass = true,
                Attributes = MemberAttributes.Public
            };
            mainClass.Members.Add(entryMethod);
            mainClass.Members.AddRange(BuildHelperMethods(bluePrints, settings, namespaces));

            // Setup the imports
            var codeNameSpaces = new List<CodeNamespaceImport>(namespaces.Count);
            foreach (var ns in namespaces)
            {
                codeNameSpaces.Add(ns.GetCodeNamespaceImport(settings.NamespaceAliasOptions));
            }
            codeNameSpaces.Sort(new CodeNamespaceImportComparer());

            mainNamespace.Imports.AddRange(codeNameSpaces.ToArray());
            mainNamespace.Types.Add(mainClass);

            // Finish up
            result.Namespaces.Add(mainNamespace);
            return result;
        }

        /// <summary>
        /// Converts an <see cref="OpenXmlPart"/> into a <see cref="string"/> representation
        /// of dotnet source code that can be compiled to build <paramref name="part"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate source code for.
        /// </param>
        /// <param name="provider">
        /// The <see cref="CodeDomProvider"/> object to create the resulting source code.
        /// </param>
        /// <returns>
        /// A <see cref="string"/> representation of the source code generated by <paramref name="provider"/>
        /// that could create <paramref name="part"/> when compiled.
        /// </returns>
        public static string GenerateSourceCode(this OpenXmlPart part, CodeDomProvider provider)
        {
            return part.GenerateSourceCode(new DefaultSerializeSettings(), provider);
        }

        /// <summary>
        /// Converts an <see cref="OpenXmlPart"/> into a <see cref="string"/> representation
        /// of dotnet source code that can be compiled to build <paramref name="part"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate source code for.
        /// </param>
        /// <param name="opts">
        /// The <see cref="NamespaceAliasOptions"/> to apply to the resulting source code.
        /// </param>
        /// <param name="provider">
        /// The <see cref="CodeDomProvider"/> object to create the resulting source code.
        /// </param>
        /// <returns>
        /// A <see cref="string"/> representation of the source code generated by <paramref name="provider"/>
        /// that could create <paramref name="part"/> when compiled.
        /// </returns>
        public static string GenerateSourceCode(this OpenXmlPart part, NamespaceAliasOptions opts, CodeDomProvider provider)
        {
            return part.GenerateSourceCode(new DefaultSerializeSettings(opts), provider);
        }

        /// <summary>
        /// Converts an <see cref="OpenXmlPart"/> into a <see cref="string"/> representation
        /// of dotnet source code that can be compiled to build <paramref name="part"/>.
        /// </summary>
        /// <param name="part">
        /// The <see cref="OpenXmlPart"/> object to generate source code for.
        /// </param>
        /// <param name="settings">
        /// The <see cref="ISerializeSettings"/> to use during the code generation
        /// process.
        /// </param>
        /// <param name="provider">
        /// The <see cref="CodeDomProvider"/> object to create the resulting source code.
        /// </param>
        /// <returns>
        /// A <see cref="string"/> representation of the source code generated by <paramref name="provider"/>
        /// that could create <paramref name="part"/> when compiled.
        /// </returns>
        public static string GenerateSourceCode(this OpenXmlPart part, ISerializeSettings settings, CodeDomProvider provider)
        {
            var codeString = new System.Text.StringBuilder();
            var code = part.GenerateSourceCode(settings);

            using (var sw = new StringWriter(codeString))
            {
                provider.GenerateCodeFromCompileUnit(code, sw,
                    new CodeGeneratorOptions() { BracingStyle = "C" });
            }
            return codeString.ToString().RemoveOutputHeaders().Trim();
        }

        #endregion

        #region Private Static Methods

        /// <summary>
        /// Checks to see if a given <see cref="OpenXmlPart"/> object requires a custom
        /// AddNewPart method to initialize.
        /// </summary>
        /// <param name="pType">
        /// The <see cref="OpenXmlPart"/> type of the object that is being added/constructed.
        /// </param>
        /// <param name="rootType">
        /// The parent <see cref="OpenXmlPart"/> type adding the <paramref name="pType"/> type.
        /// </param>
        /// <param name="addMethodName">
        /// The name of the method to use when adding <paramref name="pType"/> type to
        /// <paramref name="rootType"/> type.
        /// </param>
        /// <returns>
        /// <see langword="true"/> if <paramref name="pType"/> requires a custom AddNewPart method
        /// to initialize; otherwise <see langword="false"/>
        /// </returns>
        private static bool CheckForCustomAddNewPartMethod(Type pType, Type rootType, out string addMethodName)
        {
            var checkName = $"Add{pType.Name}";
            bool check(MethodInfo m) => m.Name.Equals(checkName, StringComparison.OrdinalIgnoreCase);

            if (rootType.GetMethods().Count(check) <= 0)
            {
                addMethodName = null;
                return false;
            }

            addMethodName = checkName;
            return true;
        }

        #endregion
    }
}