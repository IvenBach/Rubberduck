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
        private readonly string _projectId;
        private readonly RubberduckParserState _state;

        public AddComponentViewModel(RubberduckParserState state, AddComponentModel model)
            : base(model)
        {
            _state = state;

            _projectId = model.ProjectId;

            ValidateName();
            ValidateFolderPath();
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

        public bool IsValidName => !GetErrors(nameof(ComponentName))?.OfType<string>().Any() ?? true;

        private void ValidateName()
        {
            var errors = VBAIdentifierValidator.SatisfiedInvalidIdentifierCriteria(ComponentName, DeclarationType.Module).ToList();

            var nameConflicts = _state.AllUserDeclarations
                .Where(declaration => declaration.ProjectId == _projectId 
                    && declaration.ComponentName.ToUpperInvariant().Equals(Model.ComponentName.ToUpperInvariant()))
                .Any();
            if (nameConflicts)
            {
                errors.Add(string.Format(RefactoringsUI.InvalidNameCriteria_IsNotUniqueName, Model.ComponentName, Model.ProjectId));
            }

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

        public bool IsValidFolder => !GetErrors(nameof(Folder)).OfType<string>().Any();

        private void ValidateFolderPath()
        {
            if (!CodeExplorerFolderPathValidator.IsFolderPathValid(Folder, out var errors, false))
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
                base.DialogOk();
            }
        }
    }
}
