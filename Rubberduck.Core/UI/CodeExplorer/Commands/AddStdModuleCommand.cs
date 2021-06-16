using System.Collections.Generic;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Refactorings.AddComponent;
using Rubberduck.UI.CodeExplorer.AddNewComponent;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddStdModuleCommand : AddComponentCommandBase
    {
        private readonly IAddComponentService _nonCodeExplorerAddComponentService;
        private readonly IRefactoringDialogFactory _dialogFactory;

        public AddStdModuleCommand(
        ICodeExplorerAddComponentService addComponentService,
        IVbeEvents vbeEvents,
        IProjectsProvider projectsProvider,
        IAddComponentService nonCodeExplorerAddComponentService,
        IRefactoringDialogFactory dialogFactory)
            : base(addComponentService, vbeEvents, projectsProvider)
        {
            _nonCodeExplorerAddComponentService = nonCodeExplorerAddComponentService;
            _dialogFactory = dialogFactory;
        }

        public override IEnumerable<ProjectType> AllowableProjectTypes => ProjectTypes.All;

        public override ComponentType ComponentType => ComponentType.StandardModule;

        protected override void OnExecute(object parameter)
        {
            AddComponent(parameter as CodeExplorerItemViewModel);
        }

        private void AddComponent(CodeExplorerItemViewModel parameter)
        {
            ICodeExplorerNode topmostParent = parameter;
            while (topmostParent.Parent != null)
            {
                topmostParent = topmostParent.Parent;
            }
            var projectId = topmostParent.QualifiedSelection.Value.QualifiedName.ProjectId;

            var possibleFolder = ParentFolder(parameter);

            var defaultModel = new AddComponentModel("Module1", possibleFolder, projectId);

            var presenter = new AddComponentPresenter(defaultModel, _dialogFactory);
            var userEditedModel = presenter.Show();

            var folderAttribute = string.IsNullOrEmpty(userEditedModel.Folder)
                ? string.Empty
                : System.Environment.NewLine + $@"'@Folder ""{userEditedModel.Folder}""";

            var code = $@"Attribute VB_Name = ""{userEditedModel.ComponentName}""{folderAttribute}
Option Explicit

";

            _nonCodeExplorerAddComponentService.AddComponentWithAttributes(projectId, ComponentType, code, componentName: userEditedModel.ComponentName);
        }

        private string ParentFolder(object node)
        {
            if (node is CodeExplorerCustomFolderViewModel folderViewModel)
            {
                return folderViewModel.FullPath;
            }

            return node is ICodeExplorerNode codeExplorerNode && codeExplorerNode.Parent is CodeExplorerCustomFolderViewModel folderNode
                ? folderNode.FullPath
                : null;
        }
    }
}
