using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.AddComponent;

namespace Rubberduck.UI.Refactorings.AddNewComponent
{
    public partial class AddComponentView : IRefactoringView<AddComponentModel>
    {
        public AddComponentView()
        {
            InitializeComponent();

            Loaded += AddComponentView_Loaded;
        }

        private void AddComponentView_Loaded(object sender, System.Windows.RoutedEventArgs e)
        {
            componentName.Focus();
            componentName.SelectAll();

            Loaded -= AddComponentView_Loaded;
        }
    }
}
