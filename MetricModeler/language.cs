using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MetricModeler
{
    class Language
    {
        public string LanguageName { get; internal set; }
        public double Level { get; internal set; }
        public double Average { get; internal set; } // AVERAGE SOURCE STATEMENTS PER FUNCTION POINT

        public Language(string languageName, double level, double average)
        {
            LanguageName = languageName;
            Level = level;
            Average = average;
        }
        public override string ToString()
        {
            return String.Format("{0, -30}{1, -8}{2}", LanguageName, Level, Average);
        }
    }
}
