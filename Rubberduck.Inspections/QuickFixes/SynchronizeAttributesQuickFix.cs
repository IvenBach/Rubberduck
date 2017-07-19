using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public sealed class SynchronizeAttributesQuickFix : IQuickFix
    {
        private readonly RubberduckParserState _state;
        private static readonly HashSet<Type> _supportedInspections = new HashSet<Type>
        {
            typeof(MissingAnnotationInspection),
            typeof(MissingAttributeInspection),
        };

        private readonly IDictionary<string, string> _attributeNames;

        public SynchronizeAttributesQuickFix(RubberduckParserState state)
        {
            _state = state;
            _attributeNames = typeof(AnnotationType).GetFields()
                .Where(field => field.GetCustomAttributes(typeof (AttributeAnnotationAttribute), true).Any())
                .Select(a => new { AnnotationName = a.Name, a.GetCustomAttributes(typeof (AttributeAnnotationAttribute), true).Cast<AttributeAnnotationAttribute>().FirstOrDefault()?.AttributeName})
                .ToDictionary(a => a.AnnotationName, a => a.AttributeName);
        }

        public void Fix(IInspectionResult result)
        {
            var context = result.Context;
            if (result.QualifiedMemberName?.MemberName != null)
            {
                FixMember(result, context);
            }
            else
            {
                FixModule(result, context);
            }
        }

        private void FixModule(IInspectionResult result, ParserRuleContext context)
        {
            var moduleName = result.QualifiedSelection.QualifiedName;

            var attributeContext = context as VBAParser.AttributeStmtContext;
            if (attributeContext != null)
            {
                Fix(moduleName, attributeContext);
                return;
            }

            var annotationContext = context as VBAParser.AnnotationContext;
            if (annotationContext != null)
            {
                Fix(moduleName, annotationContext);
                return;
            }
        }

        private void FixMember(IInspectionResult result, ParserRuleContext context)
        {
            Debug.Assert(result.QualifiedMemberName.HasValue);
            var memberName = result.QualifiedMemberName.Value;

            var attributeContext = context as VBAParser.AttributeStmtContext;
            if (attributeContext != null)
            {
                Fix(memberName, attributeContext);
                return;
            }

            var annotationContext = context as VBAParser.AnnotationContext;
            if (annotationContext != null)
            {
                Fix(memberName, annotationContext);
                return;
            }
        }

        /// <summary>
        /// Adds an annotation to match given attribute.
        /// </summary>
        /// <param name="memberName"></param>
        /// <param name="context"></param>
        private void Fix(QualifiedMemberName memberName, VBAParser.AttributeStmtContext context)
        {
            if (context.AnnotationType() == AnnotationType.Description)
            {
                FixMemberDescriptionAnnotation(_state, memberName);
            }
            else
            {
                AddMemberAnnotation(_state, memberName, context.AnnotationType());
            }
        }

        private static void FixMemberDescriptionAnnotation(RubberduckParserState state, QualifiedMemberName memberName)
        {
            var moduleName = memberName.QualifiedModuleName;
            var rewriter = state.GetRewriter(moduleName);

            var attributes = state
                .GetModuleAttributes(moduleName)
                .Where(a => a.Key.Item1.StartsWith(memberName.MemberName)
                         && a.Key.Item2.HasFlag(DeclarationType.Member))
                .ToArray();

            Debug.Assert(attributes.Length == 1, "Member has too many attributes");
            var attribute = attributes.SingleOrDefault();

            AttributeNode node;
            if (!attribute.Value.HasMemberDescriptionAttribute(memberName.MemberName, out node))
            {
                return;
            }

            var value = node.Context.attributeValue().SingleOrDefault()?.GetText() ?? "\"\"";
            var member = state.DeclarationFinder.Members(memberName.QualifiedModuleName)
                .First(m => m.IdentifierName == memberName.MemberName);

            var insertAt = member.Context.Start;
            rewriter.InsertBefore(insertAt.TokenIndex, $"'@Description({value})\r\n");
        }

        private static void AddMemberAnnotation(RubberduckParserState state, QualifiedMemberName memberName, AnnotationType? annotationType)
        {
            Debug.Assert(annotationType.HasValue);

            var moduleName = memberName.QualifiedModuleName;
            var rewriter = state.GetRewriter(moduleName);

            var member = state.DeclarationFinder.Members(memberName.QualifiedModuleName)
                .First(m => m.IdentifierName == memberName.MemberName);

            var insertAt = member.Context.Start;
            rewriter.InsertBefore(insertAt.TokenIndex, $"'@{annotationType}\r\n");
        }

        private void Fix(QualifiedModuleName moduleName, VBAParser.AttributeStmtContext context)
        {
            
        }

        private static readonly IDictionary<AnnotationType, Action<RubberduckParserState, QualifiedModuleName>> 
            AttributeFixActions = new Dictionary<AnnotationType, Action<RubberduckParserState, QualifiedModuleName>>
            {
                [AnnotationType.PredeclaredId] = FixPredeclaredIdAttribute,
                [AnnotationType.Exposed] = FixExposedAttribute,
            }; 

        private void Fix(QualifiedModuleName moduleName, VBAParser.AnnotationContext context)
        {
            AttributeFixActions[context.AnnotationType].Invoke(_state, moduleName);
        }

        private static void FixPredeclaredIdAttribute(RubberduckParserState state, QualifiedModuleName moduleName)
        {
            var attributes = state.GetModuleAttributes(moduleName);
            var rewriter = state.GetAttributeRewriter(moduleName);
            foreach (var attribute in attributes.Values)
            {
                var predeclaredIdAttribute = attribute.PredeclaredIdAttribute;
                if (predeclaredIdAttribute == null)
                {
                    continue;
                }

                var valueToken = predeclaredIdAttribute.Context.attributeValue().Single().Start;
                Debug.Assert(valueToken.Text == "False");

                rewriter.Replace(valueToken, "True");
            }
        }

        private static void FixExposedAttribute(RubberduckParserState state, QualifiedModuleName moduleName)
        {
            var attributes = state.GetModuleAttributes(moduleName);
            var rewriter = state.GetAttributeRewriter(moduleName);
            foreach(var attribute in attributes.Values)
            {
                var exposedAttribute = attribute.ExposedAttribute;
                if(exposedAttribute == null)
                {
                    continue;
                }

                var valueToken = exposedAttribute.Context.attributeValue().Single().Start;
                Debug.Assert(valueToken.Text == "False");

                rewriter.Replace(valueToken, "True");
            }
        }

        
        /// <summary>
        /// Adds an attribute to match given annotation.
        /// </summary>
        /// <param name="memberName"></param>
        /// <param name="context"></param>
        private void Fix(QualifiedMemberName memberName, VBAParser.AnnotationContext context)
        {
            Debug.Assert(context.AnnotationType.HasFlag(AnnotationType.MemberAnnotation));

            var annotationName = Identifier.GetName(context.annotationName().unrestrictedIdentifier());
            var annotationType = context.AnnotationType;
            var attributeName =  memberName.MemberName + "." + _attributeNames[annotationName];

            var attributeInstruction = GetAttributeInstruction(context, attributeName, annotationType);
            var insertPosition = FindInsertPosition(context);

            var rewriter = _state.GetAttributeRewriter(memberName.QualifiedModuleName);
            rewriter.InsertBefore(insertPosition, attributeInstruction);
        }

        private int FindInsertPosition(VBAParser.AnnotationContext context)
        {
            return (context.AnnotatedContext as IAnnotatedContext)?.AttributeTokenIndex ?? 1;
        }

        private string GetAttributeInstruction(VBAParser.AnnotationContext context, string attributeName, AnnotationType annotationType)
        {
            string attributeInstruction = string.Empty;

            if (annotationType.HasFlag(AnnotationType.ModuleAnnotation))
            {
                switch (annotationType)
                {
                    case AnnotationType.Exposed:
                        attributeInstruction = $"Attribute {attributeName} = True\n";
                        break;
                    case AnnotationType.PredeclaredId:
                        attributeInstruction = $"Attribute {attributeName} = True\n";
                        break;
                }
            }
            else if (annotationType.HasFlag(AnnotationType.MemberAnnotation))
            {
                switch (annotationType)
                {
                    case AnnotationType.Description:
                        var description = context.annotationArgList().annotationArg().FirstOrDefault()?.GetText() ?? string.Empty;
                        description = description.StartsWith("\"") && description.EndsWith("\"")
                            ? description
                            : $"\"{description}\"";

                        attributeInstruction = $"Attribute {attributeName} = \"{description}\"\n";
                        break;
                    case AnnotationType.DefaultMember:
                        attributeInstruction = $"Attribute {attributeName} = 0";
                        break;
                    case AnnotationType.Enumerator:
                        attributeInstruction = $"Attribute {attributeName} = -4";
                        break;
                }
            }

            return attributeInstruction;
        }

        public IReadOnlyCollection<Type> SupportedInspections => _supportedInspections.ToList();

        public string Description(IInspectionResult result)
        {
            return InspectionsUI.SynchronizeAttributesQuickFix;
        }

        public bool CanFixInProcedure => false;
        public bool CanFixInModule => true;
        public bool CanFixInProject => true;
    }
}