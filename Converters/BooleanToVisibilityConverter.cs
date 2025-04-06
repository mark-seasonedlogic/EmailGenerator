using Microsoft.UI.Xaml;
using Microsoft.UI.Xaml.Data;
using System;

namespace EmailGenerator.Converters
{
    public class BooleanToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, string language)
        {
            if (value is bool boolVal)
            {
                return boolVal ? Visibility.Visible : Visibility.Collapsed;
            }
            return Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            return (value is Visibility vis && vis == Visibility.Visible);
        }
    }
}
