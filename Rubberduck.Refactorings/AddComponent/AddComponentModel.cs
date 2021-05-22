using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.AddComponent
{
    public class AddComponentModel : IRefactoringModel
    {
        public string ComponentName { get; set; }
        public string Folder { get; set; }
        public ModuleDeclaration Target { get; }
        public AddComponentModel(string componentName, string folder)
        {
            ComponentName = componentName;
            Folder = folder;
        }
    }
}
