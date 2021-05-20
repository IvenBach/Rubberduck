using System.Collections.Generic;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;


namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddStdModuleCommand : AddComponentCommandBase
    {
        private readonly IAddComponentService _nonCodeExplorerAddComponentService;

        public AddStdModuleCommand(
        ICodeExplorerAddComponentService addComponentService,
        IVbeEvents vbeEvents,
        IProjectsProvider projectsProvider,
        IAddComponentService nonCodeExplorerAddComponentService)
            : base(addComponentService, vbeEvents, projectsProvider)
        {
            _nonCodeExplorerAddComponentService = nonCodeExplorerAddComponentService;
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

            var possibleFolder = (parameter as ICodeExplorerNode).Description;

            var model = new Rubberduck.Refactorings.AddComponent.AddComponentModel("Module1", possibleFolder);

            var foo = new Rubberduck.UI.CodeExplorer.AddNewComponent.AddComponentPresenter(model, null);
            var model2 = foo.Show();

            var folderAttribute = string.IsNullOrEmpty(model2.Folder)
                ? string.Empty
                : System.Environment.NewLine + $@"'@Folder(""{model2.Folder}"")";

            var code = $@"Attribute VB_Name = ""{model2.ComponentName}""{folderAttribute}
Option Explicit

";

            _nonCodeExplorerAddComponentService.AddComponentWithAttributes(projectId, ComponentType, code, componentName: model2.ComponentName);
        }
    }
}
