using System.Collections.Generic;
using System.Text;

namespace PowerPointAddInVSTO.Extensions
{
    public static class TimingExntensions
    {
        public static string ConvertTimesToString(this List<double> timings)
        {
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < timings.Count; i++)
            {
                sb.Append(timings[i]);
                sb.Append("|");
                if (i == timings.Count - 1)
                {
                    sb.Length -= 1;
                }
            }
            sb.Insert(0, "|");
            return sb.ToString();
        }

    }
}
