using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.UI.Refactorings.AddNewComponent
{
    public class AddComponentViewModel
    {
        public string ComponentName { get; set; }
        public string Folder { get; set; }
        public AddComponentViewModel(string componentName, string folder)
        {
            ComponentName = componentName;
            Folder = folder;
        }
    }
}
