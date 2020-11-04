using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddInVSTO.ViewModel
{
    public class EffectViewModel
    {
        public string DisplayName { get; set; }
        public Slide Slide { get; set; }
        public int SlideNumber { get; set; }
        public MsoShapeType Type { get; set; }
        public float EffectTimeline { get; set; }
        public bool IsSec { get; set; }


    }
}
