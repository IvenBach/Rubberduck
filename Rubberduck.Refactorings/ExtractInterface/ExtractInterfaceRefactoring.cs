﻿using System;
using System.Linq;
using NLog;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class ExtractInterfaceRefactoring : IRefactoring
    {
        private readonly RubberduckParserState _state;
        private readonly IVBE _vbe;
        private readonly IMessageBox _messageBox;
        private readonly IRefactoringPresenterFactory _factory;
        private readonly IRewritingManager _rewritingManager;
        private ExtractInterfaceModel _model;

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public ExtractInterfaceRefactoring(RubberduckParserState state, IVBE vbe, IMessageBox messageBox, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager)
        {
            _state = state;
            _vbe = vbe;
            _rewritingManager = rewritingManager;
            _messageBox = messageBox;
            _factory = factory;
        }

        private ExtractInterfaceModel InitializeModel()
        {
            var selection = _vbe.GetActiveSelection();

            if (selection == null)
            {
                return null;
            }

            return new ExtractInterfaceModel(_state, selection.Value);
        }

        public void Refactor()
        {
            _model = InitializeModel();

            if (_model == null)
            {
                return;
            }

            using (var container = DisposalActionContainer.Create(_factory.Create<IExtractInterfacePresenter, ExtractInterfaceModel>(_model), p => _factory.Release(p)))
            {
                var presenter = container.Value;
                if (presenter == null)
                {
                    return;
                }

                _model = presenter.Show();
                if (_model == null)
                {
                    return;
                }

                using (var pane = _vbe.ActiveCodePane)
                {
                    if (pane.IsWrappingNullReference)
                    {
                        return;
                    }

                    var oldSelection = pane.GetQualifiedSelection();

                    AddInterface();

                    if (oldSelection.HasValue)
                    {
                        pane.Selection = oldSelection.Value.Selection;
                    }
                }

                _model.State.OnParseRequested(this);
            }
        }

        public void Refactor(QualifiedSelection target)
        {
            using (var pane = _vbe.ActiveCodePane)
            {
                if (pane.IsWrappingNullReference)
                {
                    return;
                }
                pane.Selection = target.Selection;
            }
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            using (var pane = _vbe.ActiveCodePane)
            {
                if (pane.IsWrappingNullReference)
                {
                    return;
                }
                pane.Selection = target.QualifiedSelection.Selection;
            }
            Refactor();
        }

        private void AddInterface()
        {
            //We need to suspend here since adding the interface and rewriting will both trigger a reparse.
            var suspendResult = _model.State.OnSuspendParser(this, new[] {ParserState.Ready}, AddInterfaceInternal);
            if (suspendResult != SuspensionResult.Completed)
            {
                _logger.Warn("Extract interface failed.");
            }
        }

        private void AddInterfaceInternal()
        {
            var targetProject = _model.TargetDeclaration.Project;
            if (targetProject == null)
            {
                return; //The target project is not available.
            }

            AddInterfaceClass(_model.TargetDeclaration, _model.InterfaceName, GetInterfaceModuleBody());

            var rewriteSession = _rewritingManager.CheckOutCodePaneSession();
            var rewriter = rewriteSession.CheckOutModuleRewriter(_model.TargetDeclaration.QualifiedModuleName);

            var firstNonFieldMember = _model.State.DeclarationFinder.Members(_model.TargetDeclaration)
                                            .OrderBy(o => o.Selection)
                                            .First(m => ExtractInterfaceModel.MemberTypes.Contains(m.DeclarationType));
            rewriter.InsertBefore(firstNonFieldMember.Context.Start.TokenIndex, $"Implements {_model.InterfaceName}{Environment.NewLine}{Environment.NewLine}");

            AddInterfaceMembersToClass(rewriter);

            rewriteSession.TryRewrite();
        }

        private void AddInterfaceClass(Declaration implementingClass, string interfaceName, string interfaceBody)
        {
            var targetProject = implementingClass.Project;
            using (var components = targetProject.VBComponents)
            {
                using (var interfaceComponent = components.Add(ComponentType.ClassModule))
                {
                    using (var interfaceModule = interfaceComponent.CodeModule)
                    {
                        interfaceComponent.Name = interfaceName;

                        var optionPresent = interfaceModule.CountOfLines > 1;
                        if (!optionPresent)
                        {
                            interfaceModule.InsertLines(1, $"{Tokens.Option} {Tokens.Explicit}{Environment.NewLine}");
                        }
                        interfaceModule.InsertLines(3, interfaceBody);
                    }
                }
            }
        }

        private void AddInterfaceMembersToClass(IModuleRewriter rewriter)
        {
            var implementInterfaceRefactoring = new ImplementInterfaceRefactoring(_vbe, _model.State, _messageBox, _rewritingManager);
            implementInterfaceRefactoring.Refactor(_model.SelectedMembers.Select(m => m.Member).ToList(), rewriter, _model.InterfaceName);
        }

        private string GetInterfaceModuleBody()
        {
            return string.Join(Environment.NewLine, _model.SelectedMembers.Select(m => m.Body));
        }
    }
}
