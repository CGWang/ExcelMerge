using System.Collections.Generic;
using System.Text;
using System.IO;

namespace ExcelMerge
{
    public class TsvReader
    {
        internal static IEnumerable<ExcelRow> Read(string path)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var sr = new StreamReader(fs, Encoding.UTF8))
            {
                var rowIndex = 0;
                while (!sr.EndOfStream)
                {
                    var columnIndex = 0;
                    var cells = new List<ExcelCell>();
                    foreach (var c in sr.ReadLine().Split('\t'))
                        cells.Add(new ExcelCell(c, columnIndex, rowIndex));

                    yield return new ExcelRow(rowIndex++, cells);
                }
            }
        }
    }
}
