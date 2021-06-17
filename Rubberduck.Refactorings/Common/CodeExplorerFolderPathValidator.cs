using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.InternalApi.Extensions;
using Rubberduck.Refactorings;

namespace Rubberduck.Refactorings.Common
{
    public static class CodeExplorerFolderPathValidator
    {
        public static bool IsFolderPathValid(string folderPath, out List<string> errors, bool treatEmptyOrNullFolderPathAsError = true)
        {
            errors = new List<string>();

            if (!treatEmptyOrNullFolderPathAsError && string.IsNullOrEmpty(folderPath))
            {
                return true;
            }

            if (treatEmptyOrNullFolderPathAsError && string.IsNullOrEmpty(folderPath))
            {
                errors.Add(RefactoringsUI.MoveFolders_EmptyFolderName);
            }
            else
            {
                if (folderPath.Any(char.IsControl))
                {
                    errors.Add(RefactoringsUI.MoveFolders_ControlCharacter);
                }

                if (folderPath.Split(FolderExtensions.FolderDelimiter).Any(string.IsNullOrEmpty))
                {
                    errors.Add(RefactoringsUI.MoveFolders_EmptySubfolderName);
                }
            }

            return !errors.Any();
        }
    }
}
