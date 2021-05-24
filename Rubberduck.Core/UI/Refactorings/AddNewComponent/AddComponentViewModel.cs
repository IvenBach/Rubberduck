using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Interaction;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AddComponent;
using Rubberduck.Refactorings.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.Refactorings.AddNewComponent
{
    public class AddComponentViewModel : RefactoringViewModelBase<AddComponentModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IMessageBox _messageBox;

        public RubberduckParserState State { get; }

        public AddComponentViewModel(RubberduckParserState state, AddComponentModel model, IMessageBox messageBox)
            : base(model)
        {
            State = state;
            _declarationFinderProvider = state;
            _messageBox = messageBox;
        }

        public string ComponentName
        {
            get => Model.ComponentName;
            set
            {
                if (value != Model.ComponentName)
                {
                    Model.ComponentName = value;
                    ValidateName();
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(ComponentName));
                    OnPropertyChanged(nameof(HasValidInputs));
                }
            }
        }

        public bool IsValidName => !GetErrors(nameof(IsValidName)).Cast<List<string>>().Any();

        private void ValidateName()
        {
            var errors = VBAIdentifierValidator.SatisfiedInvalidIdentifierCriteria(ComponentName, DeclarationType.Module).ToList();

            if (errors.Any())
            {
                SetErrors(nameof(ComponentName), errors);
            }
            else
            {
                ClearErrors(nameof(ComponentName));
            }
        }

        public string Folder
        {
            get => Model.Folder;
            set
            {
                if (value != Model.Folder)
                {
                    Model.Folder = value;
                    ValidateFolderPath();
                    OnPropertyChanged();
                    OnPropertyChanged(nameof(IsValidFolder));
                    OnPropertyChanged(nameof(HasValidInputs));
                }
            }
        }

        public bool IsValidFolder => !GetErrors(nameof(IsValidFolder)).Cast<List<string>>().Any();

        private void ValidateFolderPath()
        {
            if (!CodeExplorerFolderPathValidator.IsFolderPathValid(Folder, out var errors))
            {
                SetErrors(nameof(Folder), errors);
            }
            else
            {
                ClearErrors(nameof(Folder));
            }
        }

        public bool HasValidInputs => !HasErrors;

        protected override void DialogOk()
        {
            if (HasErrors)
            {
                DialogCancel();
            }
            else
            {
                var conflictingDeclarations = _declarationFinderProvider.DeclarationFinder
                    .UserDeclarations(DeclarationType.Module)
                    .Where(declaration => declaration.ProjectId == Model.ProjectId
                            && declaration.IdentifierName.Equals(ComponentName)); 
                
                if (conflictingDeclarations.Any()
                    && !UserConfirmsToProceedWithConflictingName(Model.ComponentName, conflictingDeclarations.FirstOrDefault()))
                {
                    base.DialogCancel();
                }
                else
                {
                    base.DialogOk();
                }
            }
        }

        private bool UserConfirmsToProceedWithConflictingName(string componentName, Declaration conflictingDeclaration)
        {
            var counter = 1;
            var suffixedName = componentName + counter.ToString();
            while (_declarationFinderProvider.DeclarationFinder
                        .UserDeclarations(DeclarationType.Module)
                        .Where(d => d.ProjectId == Model.ProjectId
                                && d.IdentifierName.Equals(suffixedName))
                        .Any())
            {
                counter++;
                suffixedName = componentName + counter.ToString();
            }

            var message = string.Format(RefactoringsUI.AddNewComponent_ConflictingDeclarations, componentName, conflictingDeclaration.QualifiedName.ToString(), suffixedName);
            return _messageBox?.ConfirmYesNo(message, RefactoringsUI.AddNewComponent_Caption) ?? false;
        }
    }
}
