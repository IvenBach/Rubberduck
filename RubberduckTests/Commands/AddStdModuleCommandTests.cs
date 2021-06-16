using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AddComponent;
using Rubberduck.UI.Refactorings.AddNewComponent;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;
using RubberduckTests.Mocks;
using Moq;
using Rubberduck.Refactorings.AnnotateDeclaration;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Refactorings;

namespace RubberduckTests.Commands
{
    [TestFixture]
    class AddStdModuleCommandTests
    {
        [Test]
        [Category("Commands")]
        public void Standard_module_successfully_added()
        {
            var worksheetObject = ("Sheet1", string.Empty, ComponentType.DocObject);

            var vbe = MockVbeBuilder.BuildFromModules(worksheetObject).Object;

            var (state, _) = MockParser.CreateAndParseWithRewritingManager(vbe);
            var codePaneComponentSourceCodeHandler = new Rubberduck.VBEditor.SourceCodeHandling.CodeModuleComponentSourceCodeHandler();
            var addComponentBaseService = new AddComponentService(state.ProjectsProvider, codePaneComponentSourceCodeHandler, codePaneComponentSourceCodeHandler);

            var dialogMock = new Mock<IRefactoringDialog<AddComponentModel, IRefactoringView<AddComponentModel>, IRefactoringViewModel<AddComponentModel>>>();
            dialogMock.Setup(dialog => dialog.ShowDialog()).Returns(RefactoringDialogResult.Execute);
            var expectedComponent = new AddComponentModel("UpdatedName", "Path.To.Module", vbe.ActiveVBProject.ProjectId);
            var dialogFactory = new Mock<IRefactoringDialogFactory>();
            dialogFactory.Setup(factory => factory.CreateDialog
            <
                AddComponentModel,
                IRefactoringView<AddComponentModel>,
                IRefactoringViewModel<AddComponentModel>,
                IRefactoringDialog<AddComponentModel, IRefactoringView<AddComponentModel>, IRefactoringViewModel<AddComponentModel>>
            >(
                It.IsAny<DialogData>(),
                It.IsAny<AddComponentModel>(),
                It.IsAny<IRefactoringView<AddComponentModel>>(),
                It.IsAny<IRefactoringViewModel<AddComponentModel>>()
            ))
            .Returns(
                (
                    DialogData dialogData,
                    AddComponentModel addComponentModel,
                    IRefactoringView<AddComponentModel> view,
                    IRefactoringViewModel<AddComponentModel> IRefactoringViewModel)
                =>
                {
                    dialogMock.SetupGet(dialog => dialog.ViewModel.Model).Returns(() => expectedComponent);
                    return dialogMock.Object;
                }
            );

            var codeExplorerAddComponentService = new Rubberduck.UI.CodeExplorer.CodeExplorerAddComponentService(state, addComponentBaseService, vbe);
            var vbeEvents = new Mock<Rubberduck.VBEditor.Events.IVbeEvents>();
            var nonCodeExplorerAddComponentService = new AddComponentService(state.ProjectsProvider, codePaneComponentSourceCodeHandler, codePaneComponentSourceCodeHandler);
            var addStdModuleCommand = new Rubberduck.UI.CodeExplorer.Commands.AddStdModuleCommand(codeExplorerAddComponentService, vbeEvents.Object, state.ProjectsProvider, nonCodeExplorerAddComponentService, dialogFactory.Object);
            
            var qualifiedModuleName = new QualifiedModuleName(vbe.ActiveVBProject);
            var qualifiedMemberName = new QualifiedMemberName(qualifiedModuleName, "NamedMember");
            var declaration = new Declaration(qualifiedMemberName, null, null, null, null, false, false, Accessibility.Public, DeclarationType.Module, false, null);
            var declarations = new List<Declaration>();
            ICodeExplorerNode projectParentNode = new CodeExplorerProjectViewModel(declaration, ref declarations, state, vbe, state.ProjectsProvider);
            var parentNode = new CodeExplorerCustomFolderViewModel(projectParentNode, "NodeName", "Path.To.Node", vbe, ref declarations);
            addStdModuleCommand.Execute(new CodeExplorerComponentViewModel(parentNode, null, ref declarations, vbe));

            var actual = vbe.ActiveVBProject.VBComponents.Single(component => component.Name == expectedComponent.ComponentName);
            var rawCode = actual.CodeModule.Content();
            (var parseTree, var _) = Rubberduck.Parsing.VBA.Parsing.VBACodeStringParser.Parse(rawCode, t => t.startRule());

            Assert.AreEqual(2, vbe.ActiveVBProject.VBComponents.Count);            
            Assert.AreEqual(expectedComponent.ComponentName, actual.Name);
            Assert.IsTrue(parseTree.GetText().Contains(expectedComponent.Folder));
        }
    }
}
