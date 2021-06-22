using System.Collections.Generic;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.CodeExplorer.Commands
{
    public class AddClassModuleCommand : AddComponentCommandBase
    {
        private readonly IAddComponentService _nonCodeExplorerAddComponentService;
        private readonly IRefactoringDialogFactory _dialogFactory;

        public AddClassModuleCommand(
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

        public override ComponentType ComponentType => ComponentType.ClassModule;

        protected override void OnExecute(object parameter)
        {
            AddClassModule(parameter as CodeExplorerItemViewModel);
        }

        private void AddClassModule(CodeExplorerItemViewModel parameter)
        {
            (var projectId, var componentName, var code) = NewComponentCodeProvider.ComponentArguments(ComponentType, _dialogFactory, parameter);

            _nonCodeExplorerAddComponentService.AddComponentWithAttributes(projectId, ComponentType, code, componentName: componentName);
        }
    }
}
