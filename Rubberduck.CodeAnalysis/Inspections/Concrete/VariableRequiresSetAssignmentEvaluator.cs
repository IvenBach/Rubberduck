﻿using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Inspections
{
    public static class VariableRequiresSetAssignmentEvaluator
    {
        /// <summary>
        /// Determines whether the 'Set' keyword is required (whether it's present or not) for the specified identifier reference.
        /// </summary>
        /// <param name="reference">The identifier reference to analyze</param>
        /// <param name="declarationFinderProvider">The parser state</param>
        public static bool RequiresSetAssignment(IdentifierReference reference, IDeclarationFinderProvider declarationFinderProvider)
        {
            if (!reference.IsAssignment)
            {
                // reference isn't assigning its declaration; not interesting
                return false;
            }

            if (reference.IsSetAssignment)
            {
                // don't assume Set keyword is legit...
                return reference.Declaration.IsObject;
            }

            var declaration = reference.Declaration;
            if (declaration.IsArray)
            {
                // arrays don't need a Set statement... todo figure out if array items are objects
                return false;
            }

            var isObjectVariable = declaration.IsObject;
            if (!isObjectVariable && !(declaration.IsUndeclared || Tokens.Variant.Equals(declaration.AsTypeName)))
            {
                return false;
            }

            // For Each iterators are implicitly set.
            var letStmtContext = reference.Context.GetAncestor<VBAParser.LetStmtContext>();
            if (reference.Context.GetAncestor<VBAParser.ForEachStmtContext>() != null && letStmtContext == null)
            {
                return false;
            }

            if (isObjectVariable)
            {
                // get the members of the returning type, a default member could make us lie otherwise
                var classModule = declaration.AsTypeDeclaration as ClassModuleDeclaration;
                if (classModule?.DefaultMember == null)
                {
                    return true;
                }
                var parameters = (classModule.DefaultMember as IParameterizedDeclaration)?.Parameters;
                // assign declaration is an object without a default parameterless (or with all parameters optional) member - LHS needs a 'Set' keyword.
                return parameters != null && parameters.All(p => p.IsOptional);
            }

            // assigned declaration is a variant. we need to know about the RHS of the assignment.           
            if (letStmtContext == null)
            {
                // not an assignment
                return false;
            }

            var expression = letStmtContext.expression();
            if (expression == null)
            {
                Debug.Assert(false, "RHS expression is empty? What's going on here?");
                return false;
            }

            if (expression is VBAParser.NewExprContext)
            {
                // RHS expression is newing up an object reference - LHS needs a 'Set' keyword:
                return true;
            }

            var literalExpression = expression as VBAParser.LiteralExprContext;
            if (literalExpression?.literalExpression()?.literalIdentifier()?.objectLiteralIdentifier() != null)
            {
                // RHS is a 'Nothing' token - LHS needs a 'Set' keyword:
                return true;
            }
            if (literalExpression != null)
            {
                return false; // any other literal expression definitely isn't an object.
            }

            // todo resolve expression return type
            var project = Declaration.GetProjectParent(reference.ParentScoping);
            var module = Declaration.GetModuleParent(reference.ParentScoping);

            //Covers the case of a single variable on the RHS of the assignment.
            var simpleName = expression.GetDescendent<VBAParser.SimpleNameExprContext>();
            if (simpleName != null && simpleName.GetText() == expression.GetText())
            {
                return declarationFinderProvider.DeclarationFinder.MatchName(simpleName.identifier().GetText())
                    .Any(d => AccessibilityCheck.IsAccessible(project, module, reference.ParentScoping, d) && d.IsObject);
            }

            // is the reference referring to something else in scope that's a object?
            return declarationFinderProvider.DeclarationFinder.MatchName(expression.GetText())
                .Any(decl => (decl.DeclarationType.HasFlag(DeclarationType.ClassModule) || Tokens.Object.Equals(decl.AsTypeName))
                && AccessibilityCheck.IsAccessible(project, module, reference.ParentScoping, decl));
        }
    }
}
