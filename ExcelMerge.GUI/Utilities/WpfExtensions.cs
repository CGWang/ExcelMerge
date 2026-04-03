using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace ExcelMerge.GUI.Utilities
{
    public static class WpfExtensions
    {
        public static void ClearChildren<T>(this Panel panel, IEnumerable<UIElement> except = null)
            where T : UIElement
        {
            var exceptSet = except != null ? new HashSet<UIElement>(except) : new HashSet<UIElement>();
            var toRemove = panel.Children.OfType<T>()
                .Where(child => !exceptSet.Contains(child))
                .Cast<UIElement>()
                .ToList();

            foreach (var child in toRemove)
                panel.Children.Remove(child);
        }

        public static int? GetRow(this Grid grid, Point point)
        {
            double y = 0;
            for (int i = 0; i < grid.RowDefinitions.Count; i++)
            {
                y += grid.RowDefinitions[i].ActualHeight;
                if (point.Y < y)
                    return i;
            }
            return null;
        }

        public static int? GetColumn(this Grid grid, Point point)
        {
            double x = 0;
            for (int i = 0; i < grid.ColumnDefinitions.Count; i++)
            {
                x += grid.ColumnDefinitions[i].ActualWidth;
                if (point.X < x)
                    return i;
            }
            return null;
        }
    }
}
