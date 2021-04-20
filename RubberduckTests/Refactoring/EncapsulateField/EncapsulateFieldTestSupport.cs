using Castle.Windsor;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;
using System;
using System.Collections.Generic;
using System.Linq;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    public class EncapsulateFieldTestSupport : EncapsulateFieldInteractiveRefactoringTest
    {
        public EncapsulateFieldTestsResolver SetupResolver(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager = null, ISelectionService selectionService = null, IIndenter indenter = null)
        {
            return GetResolver(declarationFinderProvider, rewritingManager, selectionService);
        }

        public static EncapsulateFieldTestsResolver GetResolver(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager = null, ISelectionService selectionService = null, IIndenter indenter = null)
        {
            if (declarationFinderProvider is null)
            {
                throw new ArgumentNullException("declarationFinderProvider is null");
            }

            var resolver = new EncapsulateFieldTestsResolver(declarationFinderProvider, rewritingManager, selectionService, indenter);
            resolver.Install(new WindsorContainer(), null);
            return resolver;
        }

        public string RHSIdentifier => Rubberduck.Resources.Refactorings.Refactorings.CodeBuilder_DefaultPropertyRHSParam;

        public string StateUDTDefaultTypeName => $"T{MockVbeBuilder.TestModuleName}";

        private TestEncapsulationAttributes UserModifiedEncapsulationAttributes(string field, string property = null, bool isReadonly = false, bool encapsulateFlag = true)
        {
            var testAttrs = new TestEncapsulationAttributes(field, encapsulateFlag, isReadonly);
            if (property != null)
            {
                testAttrs.PropertyName = property;
            }
            return testAttrs;
        }

        public Func<EncapsulateFieldModel, EncapsulateFieldModel> UserAcceptsDefaults(bool convertFieldToUDTMember = false)
        {
            return model =>
            {
                model.EncapsulateFieldStrategy = convertFieldToUDTMember
                    ? EncapsulateFieldStrategy.ConvertFieldsToUDTMembers
                    : EncapsulateFieldStrategy.UseBackingFields;
                return model;
            };
        }

        public Func<EncapsulateFieldModel, EncapsulateFieldModel> UserAcceptsDefaults(params string[] fieldNames)
        {
            return model =>
            {
                foreach (var name in fieldNames)
                {
                    model[name].EncapsulateFlag = true;
                }
                return model;
            };
        }

        public Func<EncapsulateFieldModel, EncapsulateFieldModel> SetParametersForSingleTarget(string field, string property = null, bool isReadonly = false, bool encapsulateFlag = true, bool asUDT = false)
        {
            var clientAttrs = UserModifiedEncapsulationAttributes(field, property, isReadonly, encapsulateFlag);

            return SetParameters(field, clientAttrs, asUDT);
        }

        public Func<EncapsulateFieldModel, EncapsulateFieldModel> SetParameters(UserInputDataObject userInput)
        {
            return model =>
            {
                if (userInput.ConvertFieldsToUDTMembers)
                {
                    model.EncapsulateFieldStrategy = EncapsulateFieldStrategy.ConvertFieldsToUDTMembers;
                    var stateUDT = model.ObjectStateUDTCandidates.Where(os => os.IdentifierName == userInput.ObjectStateUDTTargetID)
                        .Select(sfc => sfc).SingleOrDefault();

                    if (stateUDT != null)
                    {
                        model.ObjectStateUDTField = stateUDT;
                    }
                }
                else
                {
                    model.EncapsulateFieldStrategy = EncapsulateFieldStrategy.UseBackingFields;
                }

                foreach (var testModifiedAttribute in userInput.EncapsulateFieldAttributes)
                {
                    var attrsInitializedByTheRefactoring = model[testModifiedAttribute.TargetFieldName];

                    attrsInitializedByTheRefactoring.EncapsulateFlag = testModifiedAttribute.EncapsulateFlag;
                    attrsInitializedByTheRefactoring.PropertyIdentifier = testModifiedAttribute.PropertyName;
                    attrsInitializedByTheRefactoring.IsReadOnly = testModifiedAttribute.IsReadOnly;
                }
                return model;
            };
        }

        public Func<EncapsulateFieldModel, EncapsulateFieldModel> SetParameters(string originalField, TestEncapsulationAttributes attrs, bool convertFieldsToUDTMembers = false)
        {
            return model =>
            {
                model.EncapsulateFieldStrategy = convertFieldsToUDTMembers
                    ? EncapsulateFieldStrategy.ConvertFieldsToUDTMembers
                    : EncapsulateFieldStrategy.UseBackingFields;

                var encapsulatedField = model[originalField];
                encapsulatedField.EncapsulateFlag = attrs.EncapsulateFlag;
                encapsulatedField.PropertyIdentifier = attrs.PropertyName;
                encapsulatedField.IsReadOnly = attrs.IsReadOnly;
                return model;
            };
        }

        public string RefactoredCode(CodeString codeString, Func<EncapsulateFieldModel, EncapsulateFieldModel> presenterAdjustment, Type expectedException = null, bool executeViaActiveSelection = false)
            => RefactoredCode(codeString.Code, codeString.CaretPosition.ToOneBased(), presenterAdjustment, expectedException, executeViaActiveSelection);

        public IRefactoring SupportTestRefactoring(
            IRewritingManager rewritingManager,
            RubberduckParserState state,
            RefactoringUserInteraction<IEncapsulateFieldPresenter, EncapsulateFieldModel> userInteraction,
            ISelectionService selectionService)
        {
            var resolver = SetupResolver(state, rewritingManager, selectionService);
            return new EncapsulateFieldRefactoring(resolver.Resolve<EncapsulateFieldRefactoringAction>(),
                resolver.Resolve<EncapsulateFieldPreviewProvider>(),
                resolver.Resolve<IEncapsulateFieldModelFactory>(),
                userInteraction,
                selectionService,
                resolver.Resolve<ISelectedDeclarationProvider>());
        }

        public IDictionary<string, string> RefactoredCode(
            Func<EncapsulateFieldModel, EncapsulateFieldModel> presenterAction,
            TestCodeString codeString,
            params (string, string, ComponentType)[] otherModules)
        {
            return RefactoredCode(presenterAction,
                (MockVbeBuilder.TestModuleName, codeString, ComponentType.StandardModule),
                otherModules);
        }

        public IDictionary<string, string> RefactoredCode(
            Func<EncapsulateFieldModel, EncapsulateFieldModel> presenterAction,
            (string selectedModuleName, TestCodeString codeString, ComponentType componentType) moduleUnderTest,
            params (string, string, ComponentType)[] otherModules)
        {
            var modules = otherModules.ToList();

            modules.Add((moduleUnderTest.selectedModuleName, moduleUnderTest.codeString.Code, moduleUnderTest.componentType));

            return RefactoredCode(
                moduleUnderTest.selectedModuleName,
                moduleUnderTest.codeString.CaretPosition.ToOneBased(),
                presenterAction,
                null,
                false,
                modules.ToArray());
        }

        public IEncapsulateFieldCandidate RetrieveEncapsulateFieldCandidate(string inputCode, string fieldName)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            return RetrieveEncapsulateFieldCandidate(vbe, fieldName);
        }

        public IEncapsulateFieldCandidate RetrieveEncapsulateFieldCandidate(string inputCode, string fieldName, DeclarationType declarationType)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            return RetrieveEncapsulateFieldCandidate(vbe, fieldName, declarationType);
        }

        public IEncapsulateFieldCandidate RetrieveEncapsulateFieldCandidate(IVBE vbe, string fieldName, DeclarationType declarationType = DeclarationType.Variable)
        {
            var state = MockParser.CreateAndParse(vbe);
            using (state)
            {
                var resolver = SetupResolver(state);

                var match = state.DeclarationFinder.MatchName(fieldName).Where(m => m.DeclarationType.Equals(declarationType)).Single();

                var model = resolver.Resolve<IEncapsulateFieldModelFactory>().Create(match);

                model.ConflictFinder.AssignNoConflictIdentifiers(model[match.IdentifierName]);

                return model[match.IdentifierName];
            }
        }

        public EncapsulateFieldModel RetrieveUserModifiedModelPriorToRefactoring(string inputCode, string declarationName, DeclarationType declarationType, Func<EncapsulateFieldModel, EncapsulateFieldModel> presenterAdjustment)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _).Object;
            return RetrieveUserModifiedModelPriorToRefactoring(vbe, declarationName, declarationType, presenterAdjustment);
        }

        public EncapsulateFieldModel RetrieveUserModifiedModelPriorToRefactoring(IVBE vbe, string declarationName, DeclarationType declarationType, Func<EncapsulateFieldModel, EncapsulateFieldModel> presenterAdjustment)
        {
            var initialModel = InitialModel(vbe, declarationName, declarationType);
            return presenterAdjustment(initialModel);
        }

        protected override IRefactoring TestRefactoring(
            IRewritingManager rewritingManager,
            RubberduckParserState state,
            RefactoringUserInteraction<IEncapsulateFieldPresenter, EncapsulateFieldModel> userInteraction,
            ISelectionService selectionService)
        {
            return SupportTestRefactoring(rewritingManager, state, userInteraction, selectionService);
        }

        public string RefactoredCode<TRefactoring,TModel>(string code, Func<RubberduckParserState, EncapsulateFieldTestsResolver, TModel> modelBuilder, IIndenter indenter = null) where TRefactoring : CodeOnlyRefactoringActionBase<TModel> where TModel : class, IRefactoringModel
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(code, out _).Object;
            var componentName = vbe.SelectedVBComponent.Name;
            var refactored = RefactoredCode<TRefactoring,TModel>(vbe, modelBuilder, indenter);
            return refactored[componentName];
        }

        public IDictionary<string, string> RefactoredCode<TRefactoring,TModel>(IVBE vbe, Func<RubberduckParserState, EncapsulateFieldTestsResolver, TModel> modelBuilder, IIndenter indenter = null) where TRefactoring: CodeOnlyRefactoringActionBase<TModel> where TModel: class, IRefactoringModel
        {
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                
                var resolver = GetResolver(state, rewritingManager, indenter: indenter);

                var refactoring = resolver.Resolve<TRefactoring>();

                var model = modelBuilder(state, resolver);

                refactoring.Refactor(model);

                return vbe.ActiveVBProject.VBComponents
                    .ToDictionary(component => component.Name, component => component.CodeModule.Content());
            }
        }
    }

    public class TestEncapsulationAttributes
    {
        public TestEncapsulationAttributes(string fieldName, bool encapsulationFlag = true, bool isReadOnly = false)
        {
            _identifiers = new EncapsulationIdentifiers(fieldName);
            EncapsulateFlag = encapsulationFlag;
            IsReadOnly = isReadOnly;
        }

        private EncapsulationIdentifiers _identifiers;
        public string TargetFieldName => _identifiers.TargetFieldName;

        public string NewFieldName
        {
            get => _identifiers.Field;
            set => _identifiers.Field = value;
        }
        public string PropertyName
        {
            get => _identifiers.Property;
            set => _identifiers.Property = value;
        }
        public bool EncapsulateFlag { get; set; }
        public bool IsReadOnly { get; set; }
    }

    public class UserInputDataObject
    {
        private List<TestEncapsulationAttributes> _userInput = new List<TestEncapsulationAttributes>();
        private List<(string, string, bool)> _udtNameFlagPairs = new List<(string, string, bool)>();

        public UserInputDataObject() { }

        public UserInputDataObject UserSelectsField(string fieldName, string propertyName = null, bool isReadOnly = false)
        {
            return AddUserInputSet(fieldName, propertyName, true, isReadOnly);
        }

        public UserInputDataObject AddUserInputSet(string fieldName, string propertyName = null, bool encapsulationFlag = true, bool isReadOnly = false)
        {
            var attrs = new TestEncapsulationAttributes(fieldName, encapsulationFlag, isReadOnly);
            attrs.PropertyName = propertyName ?? attrs.PropertyName;
            attrs.EncapsulateFlag = encapsulationFlag;
            attrs.IsReadOnly = isReadOnly;

            _userInput.Add(attrs);
            return this;
        }

        public bool ConvertFieldsToUDTMembers { set; get; }

        public void EncapsulateUsingUDTField(string targetID = null)
        {
            ObjectStateUDTTargetID = targetID;
            ConvertFieldsToUDTMembers = true;
        }

        public string ObjectStateUDTTargetID { set; get; }

        public string StateUDT_TypeName { set; get; }

        public string StateUDT_FieldName { set; get; }

        public TestEncapsulationAttributes this[string fieldName]
            => EncapsulateFieldAttributes.Where(efa => efa.TargetFieldName == fieldName).Single();

        public IEnumerable<TestEncapsulationAttributes> EncapsulateFieldAttributes => _userInput;
    }
}