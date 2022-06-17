using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace StegoLine.Utils.Converters {
    [ValueConversion(typeof(int), typeof(double))]
    public class WidthConverter: IValueConverter {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture) {
            double Width = System.Convert.ToDouble(value);
            return Width * 0.05 / 635.0;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) {
            double Width = System.Convert.ToDouble(value, culture.NumberFormat);
            double iWidth = System.Convert.ToInt32(Width * 635.0 / 0.05);
            return iWidth;
        }
    }
}
