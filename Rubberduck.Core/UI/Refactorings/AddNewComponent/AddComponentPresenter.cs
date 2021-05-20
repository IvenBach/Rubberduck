using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AddComponent;
using Rubberduck.UI.Refactorings;


namespace Rubberduck.UI.CodeExplorer.AddNewComponent
{
    class AddComponentPresenter : RefactoringPresenterBase<AddComponentModel>
    {
        private static readonly DialogData DialogData = DialogData.Create(RefactoringsUI.AddNewComponent_Caption, 156, 400);
        public AddComponentPresenter(AddComponentModel model, IRefactoringDialogFactory factory)
            : base(DialogData ,model, factory)
        {}
    }
}
