using Autodesk.Navisworks.Api;
using System;
using System.Collections.Generic;
using System.Linq;

namespace NwUtils.ProjectTree
{
    public static class TreeFunctions
    {
        public static ModelItemCollection Select(ModelItem startPoint, Func<ModelItem, bool> filter, string description = "")
        {
            ModelItemCollection result = new ModelItemCollection();
            var stack = new Stack<ModelItem>();
            stack.Push(startPoint);
            var progress = Application.BeginProgress(description);

            int processed = 0;
            int estimatedTotal = 1;

            try
            {
                while (stack.Count > 0)
                {
                    if (progress.IsCanceled)
                        break;

                    var current = stack.Pop();
                    processed++;

                    if (filter(current))
                        result.Add(current);

                    foreach (var child in current.Children)
                    {
                        stack.Push(child);
                        estimatedTotal++;
                    }

                    double fraction = Math.Min(1.0, (double)processed / estimatedTotal);
                    progress.Update(fraction);
                }
            }
            finally
            {
                Application.EndProgress();
            }

            return result;
        }

        public static void Where(IEnumerable<ModelItem> startPoints, Func<ModelItem, bool> condition, Func<ModelItem, bool> action, string description = "")
        {
            var stack = new Stack<ModelItem>(startPoints);
            var progress = Application.BeginProgress(description);
            int processed = 0;
            int estimatedTotal = stack.Count;

            try
            {
                while (stack.Count > 0)
                {
                    if (progress.IsCanceled)
                        break;

                    var current = stack.Pop();
                    processed++;

                    if (condition(current))
                    {
                        if (!action(current))
                            break;
                    }

                    if (current.Children != null && current.Children.Any())
                    {
                        foreach (var child in current.Children)
                        {
                            if (progress.IsCanceled)
                                break;

                            stack.Push(child);
                            estimatedTotal++;
                        }
                    }

                    double fraction = Math.Min(1.0, (double)processed / estimatedTotal);
                    progress.Update(fraction);
                }
            }
            finally
            {
                Application.EndProgress();
                stack.Clear();
            }
        }

        public static void ForEach(SavedItem set, Action<ModelItem> action, string description = "")
        {
            var items = ((SelectionSet)set).GetSelectedItems();
            int total = items.Count;
            int processed = 0;

            var progress = Application.BeginProgress(description);
            try
            {
                foreach (var item in items)
                {
                    if (progress.IsCanceled)
                        break;

                    processed++;
                    double progressFraction = (double)processed / total;
                    progress.Update(progressFraction);

                    if (item is ModelItem modelItem)
                        action(modelItem);
                }
            }
            finally
            {
                Application.EndProgress();
            }
        }
    }
}
