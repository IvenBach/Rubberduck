using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.AddComponent
{
    public class AddComponentModel
    {
        public string ComponentName { get; set; }
        public string Folder { get; set; }
        public AddComponentModel(string componentName, string folder)
        {
            ComponentName = componentName;
            Folder = folder;
        }
    }
}
