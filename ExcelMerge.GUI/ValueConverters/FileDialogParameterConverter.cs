using System;
using System.Globalization;
using System.Linq;
using System.Windows;
using System.Windows.Data;

namespace ExcelMerge.GUI.ValueConverters
{
    public class FileDialogParameterConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            if (values == null || values.Length < 2 ||
                values[0] == null || values[0] == DependencyProperty.UnsetValue ||
                values[1] == null || values[1] == DependencyProperty.UnsetValue)
                return null;

            var obj = values[0];
            var propertyName = values[1] as string;
            var propertyInfo = obj.GetType().GetProperties().FirstOrDefault(p => p.Name == propertyName);

            return new FileDialogParameter(obj, propertyInfo);
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }
}
