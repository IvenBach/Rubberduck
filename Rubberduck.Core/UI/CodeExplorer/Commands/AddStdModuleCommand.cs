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
            var folderAttribute = string.IsNullOrEmpty(possibleFolder)
                ? string.Empty
                : System.Environment.NewLine + $@"'@Folder(""{possibleFolder}"")";
            
            //prompt for ComponentName & Annotations
            var componentName = "RenamedStandardModule";
            var code = $@"Attribute VB_Name = ""{componentName}""{folderAttribute}
Option Explicit

";
            
            _nonCodeExplorerAddComponentService.AddComponentWithAttributes(projectId, ComponentType, code, componentName: "NonDefaultName");
        }
    }
}
