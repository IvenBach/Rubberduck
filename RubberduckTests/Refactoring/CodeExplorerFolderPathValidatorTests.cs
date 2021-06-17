using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    class CodeExplorerFolderPathValidatorTests
    {
        [Test]
        [Category("Refactoring")]
        [TestCase("", true, false)]
        [TestCase("", false, true)]
        [TestCase(null, true, false)]
        [TestCase(null, false, true)]
        public void Empty_or_null_string_path_is_treated_as_error(string folderPath, bool treatEmptyOrNullAsError, bool expected)
        {
            var actual = Rubberduck.Refactorings.Common.CodeExplorerFolderPathValidator.IsFolderPathValid(folderPath, out var _, treatEmptyOrNullAsError);
            Assert.AreEqual(expected, actual);
        }

        [Test]
        [Category("Refactoring")]
        [TestCase("Path.with.\0.control.chars")]
        [TestCase("Path.with.\r\n.control.chars")]
        public void Path_containing_control_character_is_not_valid(string path)
        {
            var actual = Rubberduck.Refactorings.Common.CodeExplorerFolderPathValidator.IsFolderPathValid(path, out var _);

            Assert.IsFalse(actual);
        }

        [Test]
        [Category("Refactoring")]
        [TestCase(".Path.to.module")]
        [TestCase("Path..module")]
        [TestCase("Path.to.module.")]
        public void Empty_folder_or_sub_folder_is_not_valid(string path)
        {
            var actual = Rubberduck.Refactorings.Common.CodeExplorerFolderPathValidator.IsFolderPathValid(path, out var _);

            Assert.IsFalse(actual);
        }
    }
}
