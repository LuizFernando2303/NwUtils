using Autodesk.Navisworks.Api;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NwUtils.Utils
{
    public static class UtilsFunctions
    {
        public static List<SavedItem> GetSelectionSets(SavedItem root = null)
        {
            if (root == null)
                root = Application.ActiveDocument.SelectionSets.RootItem;
            var result = new List<SavedItem>();

            if (root == null)
                return result;

            if (root.IsGroup && root is GroupItem group)
            {
                foreach (var child in group.Children)
                {
                    result.AddRange(GetSelectionSets(child));
                }
            }

            if (!string.IsNullOrWhiteSpace(root.DisplayName))
            {
                result.Add(root);
            }

            return result;
        }
    }
}
