using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Navigation.CodeExplorer;
using Rubberduck.Refactorings.AddComponent;
using Rubberduck.UI.CodeExplorer.AddNewComponent;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.CodeExplorer
{
    public static class NewComponentCodeProvider
    {
        public static (string ProjectId, string ComponentName, string Code) ComponentArguments(ComponentType componentType, IRefactoringDialogFactory factory, object parameter)
        {
            var viewModel = parameter as CodeExplorerItemViewModel;
            var defaultName = DefaultName(componentType);
            var projectId = ProjectId(viewModel);
            var possibleFolder = ParentFolder(viewModel);            
            var defaultModel = new AddComponentModel(defaultName, possibleFolder, projectId);

            var presenter = new AddComponentPresenter(defaultModel, factory);
            var userEditedModel = presenter.Show();

            var code = RawCode(componentType, userEditedModel);

            return (projectId, userEditedModel.ComponentName, code);
        }

        private static string ProjectId(ICodeExplorerNode node)
        {
            ICodeExplorerNode topmostParent = node;
            while (topmostParent.Parent != null)
            {
                topmostParent = topmostParent.Parent;
            }

            return topmostParent.QualifiedSelection.Value.QualifiedName.ProjectId;
        }

        private static string ParentFolder(object node)
        {
            if (node is CodeExplorerCustomFolderViewModel folderViewModel)
            {
                return folderViewModel.FullPath;
            }

            return node is ICodeExplorerNode codeExplorerNode && codeExplorerNode.Parent is CodeExplorerCustomFolderViewModel folderNode
                ? folderNode.FullPath
                : null;
        }

        private static string DefaultName(ComponentType componentType)
        {
            switch (componentType)
            {
                case ComponentType.StandardModule:
                    return "Module1";
                case ComponentType.ClassModule:
                    return "Class1";
                case ComponentType.UserForm:
                    return "Userform1";
                //case ComponentType.ResFile:
                //    break;
                //case ComponentType.VBForm:
                //    break;
                //case ComponentType.MDIForm:
                //    break;
                //case ComponentType.PropPage:
                //    break;
                //case ComponentType.UserControl:
                //    break;
                //case ComponentType.DocObject:
                //    break;
                //case ComponentType.RelatedDocument:
                //    break;
                //case ComponentType.ActiveXDesigner:
                //    break;
                //case ComponentType.Document:
                //    break;
                //case ComponentType.ComComponent:
                //    break;
                //case ComponentType.Undefined:
                //    break;
                default:
                    return null;
            }
        }

        private static string RawCode(ComponentType componentType, AddComponentModel model)
        {
            var folderAttribute = string.IsNullOrEmpty(model.Folder)
                ? string.Empty
                : Environment.NewLine + $@"'@Folder ""{model.Folder}""";
            string code = null;
            switch (componentType)
            {
                case ComponentType.StandardModule:
                    code = $@"Attribute VB_Name = ""{model.ComponentName}""{folderAttribute}
Option Explicit

";
                    break;
                case ComponentType.ClassModule:
                    code = $@"VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = ""{ model.ComponentName}""
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False{folderAttribute}
Option Explicit

";
                    break;
                //case ComponentType.UserForm:
                //    break;
                //case ComponentType.ResFile:
                //    break;
                //case ComponentType.VBForm:
                //    break;
                //case ComponentType.MDIForm:
                //    break;
                //case ComponentType.PropPage:
                //    break;
                //case ComponentType.UserControl:
                //    break;
                //case ComponentType.DocObject:
                //    break;
                //case ComponentType.RelatedDocument:
                //    break;
                //case ComponentType.ActiveXDesigner:
                //    break;
                //case ComponentType.Document:
                //    break;
                //case ComponentType.ComComponent:
                //    break;
                //case ComponentType.Undefined:
                //    break;
            }

            return code;
        }
    }
}
