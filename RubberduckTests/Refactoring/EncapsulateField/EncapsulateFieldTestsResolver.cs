﻿using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.EncapsulateField;
using Rubberduck.Refactorings.ReplaceReferences;
using Rubberduck.Refactorings.ReplacePrivateUDTMemberReferences;
using Rubberduck.Refactorings.ReplaceDeclarationIdentifier;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingUDTMember;
using Rubberduck.Refactorings.EncapsulateFieldUseBackingField;
using Rubberduck.Refactorings.EncapsulateFieldInsertNewCode;
using Rubberduck.SmartIndenter;
using RubberduckTests.Settings;
using Rubberduck.Refactorings.ModifyUserDefinedType;
using Castle.Windsor;
using Castle.Facilities.TypedFactory;
using Castle.MicroKernel.Registration;
using Moq;
using System;
using Rubberduck.Parsing.UIContext;
using Rubberduck.VBEditor.Utility;
using Rubberduck.UI.Command.Refactorings;
using Rubberduck.Interaction;
using Rubberduck.UI.Command.Refactorings.Notifiers;
using Castle.MicroKernel.SubSystems.Configuration;

namespace RubberduckTests.Refactoring.EncapsulateField
{
    public class EncapsulateFieldTestsResolver : IWindsorInstaller
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IRewritingManager _rewritingManager;
        private readonly ICodeBuilder _codeBuilder;
        private readonly IIndenter _testIndenter;
        private readonly IUiDispatcher _uiDispatcher;
        private readonly IRefactoringPresenterFactory _presenterFactory;
        private readonly ISelectionService _selectionService;
        private readonly IMessageBox _messageBox;

        private IWindsorContainer _container;

        public EncapsulateFieldTestsResolver(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager = null, ISelectionService selectionService = null, IIndenter indenter = null)
        {
            _declarationFinderProvider = declarationFinderProvider;

            _rewritingManager = rewritingManager;

            _selectionService = selectionService;

            _testIndenter = indenter ?? new Indenter(null, () =>
               {
                   var s = IndenterSettingsTests.GetMockIndenterSettings();
                   s.VerticallySpaceProcedures = true;
                   s.LinesBetweenProcedures = 1;
                   return s;
               });

            _codeBuilder = new CodeBuilder(_testIndenter);

            _presenterFactory = new Mock<IRefactoringPresenterFactory>().Object;

            var uiDispatcherMock = new Mock<IUiDispatcher>();
            uiDispatcherMock
                .Setup(m => m.Invoke(It.IsAny<Action>()))
                .Callback((Action action) => action.Invoke());

            _uiDispatcher = uiDispatcherMock.Object;

            _messageBox = new Mock<IMessageBox>().Object;
        }

        public void Install(IWindsorContainer container, IConfigurationStore store) 
            => Install(container);

        public T Resolve<T>() where T : class => _container.Resolve<T>() as T;

        private void Install(IWindsorContainer container)
        {
            _container = container;
            RegisterInstances(_container);
            RegisterSingletonObjects(container);
            RegisterInterfaceToImplementationPairsSingleton(container);
            RegisterInterfaceToImplementationPairsTransient(container);
            RegisterAutoFactories(container);
        }

        private void RegisterInstances(IWindsorContainer container)
        {
            container.Kernel.Register(Component.For<IDeclarationFinderProvider, RubberduckParserState>().Instance(_declarationFinderProvider));
            container.Kernel.Register(Component.For<IIndenter>().Instance(_testIndenter));
            container.Kernel.Register(Component.For<ICodeBuilder>().Instance(_codeBuilder));
            if (_rewritingManager != null)
            {
                container.Kernel.Register(Component.For<IRewritingManager>().Instance(_rewritingManager));
            }
            if (_selectionService != null)
            {
                container.Kernel.Register(Component.For<ISelectionProvider>().Instance(_selectionService));
            }
            container.Kernel.Register(Component.For<IUiDispatcher>().Instance(_uiDispatcher));
            container.Kernel.Register(Component.For<IRefactoringPresenterFactory>().Instance(_presenterFactory));
            container.Kernel.Register(Component.For<IMessageBox>().Instance(_messageBox));
        }

        private static void RegisterSingletonObjects(IWindsorContainer container)
        {
            container.Kernel.Register(Component.For<EncapsulateFieldRefactoring>());
            container.Kernel.Register(Component.For<EncapsulateFieldRefactoringAction>());
            container.Kernel.Register(Component.For<EncapsulateFieldUseBackingUDTMemberRefactoringAction>());
            container.Kernel.Register(Component.For<EncapsulateFieldUseBackingFieldRefactoringAction>());
            container.Kernel.Register(Component.For<ReplacePrivateUDTMemberReferencesRefactoringAction>());
            container.Kernel.Register(Component.For<EncapsulateFieldInsertNewCodeRefactoringAction>());
            container.Kernel.Register(Component.For<ModifyUserDefinedTypeRefactoringAction>());
            container.Kernel.Register(Component.For<ReplaceReferencesRefactoringAction>());
            container.Kernel.Register(Component.For<ReplaceDeclarationIdentifierRefactoringAction>());
            container.Kernel.Register(Component.For<EncapsulateFieldPreviewProvider>());
            container.Kernel.Register(Component.For<EncapsulateFieldUseBackingFieldPreviewProvider>());
            container.Kernel.Register(Component.For<EncapsulateFieldUseBackingUDTMemberPreviewProvider>());
            container.Kernel.Register(Component.For<EncapsulateFieldFailedNotifier>());
            container.Kernel.Register(Component.For<RefactorEncapsulateFieldCommand>());
            container.Kernel.Register(Component.For<RefactoringUserInteraction<IEncapsulateFieldPresenter, EncapsulateFieldModel>>());
        }

        private static void RegisterInterfaceToImplementationPairsSingleton(IWindsorContainer container)
        {
            container.Kernel.Register(Component.For<ISelectedDeclarationProvider>()
                .ImplementedBy<SelectedDeclarationProvider>());

            container.Kernel.Register(Component.For<IEncapsulateFieldModelFactory>()
                .ImplementedBy<EncapsulateFieldModelFactory>());

            container.Kernel.Register(Component.For<IEncapsulateFieldUseBackingUDTMemberModelFactory>()
                .ImplementedBy<EncapsulateFieldUseBackingUDTMemberModelFactory>());

            container.Kernel.Register(Component.For<IEncapsulateFieldUseBackingFieldModelFactory>()
                .ImplementedBy<EncapsulateFieldUseBackingFieldModelFactory>());

            container.Kernel.Register(Component.For<IEncapsulateFieldCandidateFactory>()
                .ImplementedBy<EncapsulateFieldCandidateFactory>());

            container.Kernel.Register(Component.For<IPropertyAttributeSetsGenerator>()
                .ImplementedBy<PropertyAttributeSetsGenerator>());

            container.Kernel.Register(Component.For<IEncapsulateFieldCodeBuilder>()
               .ImplementedBy<EncapsulateFieldCodeBuilder>());

            container.Kernel.Register(Component.For<IEncapsulateFieldRefactoringActionsProvider>()
               .ImplementedBy<EncapsulateFieldRefactoringActionsProvider>());

            container.Kernel.Register(Component.For<IReplacePrivateUDTMemberReferencesModelFactory>()
                .ImplementedBy<ReplacePrivateUDTMemberReferencesModelFactory>());
        }

        private static void RegisterInterfaceToImplementationPairsTransient(IWindsorContainer container)
        {
            container.Kernel.Register(Component.For<INewContentAggregator>()
                .ImplementedBy<NewContentAggregator>()
                .LifestyleTransient());

            container.Kernel.Register(Component.For<IEncapsulateFieldConflictFinder>()
                .ImplementedBy<EncapsulateFieldConflictFinder>()
                .LifestyleTransient());

            container.Kernel.Register(Component.For<IEncapsulateFieldCandidateSetsProvider>()
                .ImplementedBy<EncapsulateFieldCandidateSetsProvider>()
                .LifestyleTransient());
        }

        private static void RegisterAutoFactories(IWindsorContainer container)
        {
            container.Kernel.AddFacility<TypedFactoryFacility>();
            container.Kernel.Register(Component.For<IEncapsulateFieldCandidateSetsProviderFactory>().AsFactory().LifestyleSingleton());
            container.Kernel.Register(Component.For<IEncapsulateFieldConflictFinderFactory>().AsFactory().LifestyleSingleton());
            container.Kernel.Register(Component.For<INewContentAggregatorFactory>().AsFactory().LifestyleSingleton());
        }
    }
}
