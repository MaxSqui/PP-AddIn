using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddInVSTO.ViewModel
{
    public class SlideInfoViewModel
    {
        public int Number { get; set; }
        public int Clicks { get; set; }
        public float SlideTime { get; set; }
        public Slide Slide { get; set; }
    }
}
