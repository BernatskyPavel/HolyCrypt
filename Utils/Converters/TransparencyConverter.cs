using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace StegoLine.Utils.Converters {

    [ValueConversion(typeof(double), typeof(int))]
    public class TransparencyConverter: IValueConverter {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => System.Convert.ToDouble(value) / 1000.0;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => System.Convert.ToDouble(value) * 1000.0;
    }
}
