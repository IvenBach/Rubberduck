using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.UI.Refactorings.AddNewComponent;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using NUnit.Framework;

namespace RubberduckTests.Refactoring.AddComponent
{
    [TestFixture]
    class AddComponentTests
    {
        [Test]
        [Category("Refactorings")]
        [Category("AddComponent")]
        public void AddComponentViewModel_with_conflicting_name_found_only_in_same_project()
        {
            var vbeBuilder = new MockVbeBuilder();

            string moduleContent = $"Public Sub FooMember(){Environment.NewLine}End Sub";
            var nonConflictingProject = vbeBuilder.ProjectBuilder("FirstProject", ProjectProtection.Unprotected);            
            nonConflictingProject.AddComponent("Module1", ComponentType.StandardModule, moduleContent);
            var mockedNonConflictingProject = nonConflictingProject.Build();

            var conflictingProject = vbeBuilder.ProjectBuilder("SecondProject", ProjectProtection.Unprotected);
            conflictingProject.AddComponent("Module2", ComponentType.StandardModule, moduleContent);
            const string conflictingName = "ConflictingName";
            conflictingProject.AddComponent(conflictingName, ComponentType.StandardModule, moduleContent);
            var mockedConflictingProject = conflictingProject.Build();
            
            vbeBuilder.AddProject(mockedNonConflictingProject);
            vbeBuilder.AddProject(mockedConflictingProject);

            var vbe = vbeBuilder.Build().Object;
            var state = MockParser.CreateAndParse(vbe);

            var validModel = new Rubberduck.Refactorings.AddComponent.AddComponentModel("NonConflictingName", string.Empty, mockedNonConflictingProject.Object.ProjectId);
            var validViewModel = new AddComponentViewModel(state, validModel);

            var conflictingModel = new Rubberduck.Refactorings.AddComponent.AddComponentModel(conflictingName, string.Empty, mockedConflictingProject.Object.ProjectId);
            var viewModelWithConflict = new AddComponentViewModel(state, conflictingModel);

            Assert.IsTrue(validViewModel.IsValidName);
            Assert.IsFalse(viewModelWithConflict.IsValidName);
        }
    }
}
