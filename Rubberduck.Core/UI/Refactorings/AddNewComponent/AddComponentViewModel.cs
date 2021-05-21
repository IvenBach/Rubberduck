using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Refactorings.AddComponent;

namespace Rubberduck.UI.Refactorings.AddNewComponent
{
    public class AddComponentViewModel : RefactoringViewModelBase<AddComponentModel>
    {
        public string ComponentName { get; set; }
        public string Folder { get; set; }
        public AddComponentViewModel(AddComponentModel model, string componentName, string folder)
            : base(model)
        {
            ComponentName = componentName;
            Folder = folder;
        }
    }
}
