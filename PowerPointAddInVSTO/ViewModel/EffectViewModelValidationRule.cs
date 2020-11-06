using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace PowerPointAddInVSTO.ViewModel
{
    public class EffectViewModelValidationRule : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            float currentTiming = 0;
            try
            {
                if (((float)value) > 0)
                    currentTiming = float.Parse((String)value);
            }
            catch (Exception e)
            {
                return new ValidationResult(false, "Illegal characters or " + e.Message);
            }
            return new ValidationResult(true, null);
        }
    }
}
