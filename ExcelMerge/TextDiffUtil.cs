using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NetDiff;

namespace ExcelMerge
{
    public static class TextDiffUtil
    {
        /// <summary>
        /// Computes character-level inline diff segments for the DESTINATION side.
        /// Deleted characters (present only in src) are skipped; Inserted and Equal characters are included.
        /// </summary>
        public static List<(string Text, bool IsModified)> ComputeInlineDiff(string src, string dst)
        {
            if (string.IsNullOrEmpty(src) && string.IsNullOrEmpty(dst))
                return new List<(string, bool)>();

            var results = DiffUtil.Diff(
                src?.ToCharArray() ?? Array.Empty<char>(),
                dst?.ToCharArray() ?? Array.Empty<char>());
            results = DiffUtil.Order(results, DiffOrderType.LazyDeleteFirst);
            results = DiffUtil.OptimizeCaseDeletedFirst(results);

            return BuildSegments(results, isSrc: false);
        }

        /// <summary>
        /// Computes character-level inline diff segments for the SOURCE side.
        /// Inserted characters (present only in dst) are skipped; Deleted and Equal characters are included.
        /// </summary>
        public static List<(string Text, bool IsModified)> ComputeInlineDiffSrc(string src, string dst)
        {
            if (string.IsNullOrEmpty(src) && string.IsNullOrEmpty(dst))
                return new List<(string, bool)>();

            var results = DiffUtil.Diff(
                src?.ToCharArray() ?? Array.Empty<char>(),
                dst?.ToCharArray() ?? Array.Empty<char>());
            results = DiffUtil.Order(results, DiffOrderType.LazyDeleteFirst);
            results = DiffUtil.OptimizeCaseDeletedFirst(results);

            return BuildSegments(results, isSrc: true);
        }

        private static List<(string Text, bool IsModified)> BuildSegments(
            IEnumerable<DiffResult<char>> results, bool isSrc)
        {
            var segments = new List<(string Text, bool IsModified)>();
            var currentText = new StringBuilder();
            bool? currentIsModified = null;

            foreach (var r in results)
            {
                // For src side: skip Inserted chars (they belong to dst only)
                // For dst side: skip Deleted chars (they belong to src only)
                if (isSrc && r.Status == DiffStatus.Inserted)
                    continue;
                if (!isSrc && r.Status == DiffStatus.Deleted)
                    continue;

                bool isModified = r.Status != DiffStatus.Equal;
                char ch;

                if (r.Status == DiffStatus.Equal)
                    ch = r.Obj1;
                else if (r.Status == DiffStatus.Modified)
                    ch = isSrc ? r.Obj1 : r.Obj2;
                else if (r.Status == DiffStatus.Deleted)
                    ch = r.Obj1;
                else // Inserted
                    ch = r.Obj2;

                if (currentIsModified.HasValue && currentIsModified.Value != isModified)
                {
                    segments.Add((currentText.ToString(), currentIsModified.Value));
                    currentText.Clear();
                }
                currentIsModified = isModified;
                currentText.Append(ch);
            }

            if (currentText.Length > 0 && currentIsModified.HasValue)
                segments.Add((currentText.ToString(), currentIsModified.Value));

            return segments;
        }
    }
}
