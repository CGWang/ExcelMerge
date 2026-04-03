using System;
using System.Collections.Generic;
using System.Linq;

namespace ExcelMerge.GUI.Utilities
{
    public static class CollectionExtensions
    {
        public static IEnumerable<IEnumerable<T>> SplitByRegularity<T>(
            this IEnumerable<T> source,
            Func<IEnumerable<T>, T, bool> predicate)
        {
            var group = new List<T>();

            foreach (var item in source)
            {
                if (group.Count > 0 && !predicate(group, item))
                {
                    yield return group;
                    group = new List<T>();
                }
                group.Add(item);
            }

            if (group.Count > 0)
                yield return group;
        }
    }
}
