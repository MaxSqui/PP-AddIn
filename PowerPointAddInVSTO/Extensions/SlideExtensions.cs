using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddInVSTO.Extensions
{
    public static class SlideExtensions
    {
        public static Shape GetAudioShape(this Slide slide)
        {
            foreach (Shape shape in slide.Shapes)
            {
                if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoMedia)
                {
                    return shape;
                }
            }
            return null;
        }

        public static void RemoveAnimationTrigger(this Slide slide, Shape shape)
        {
            foreach(Sequence sequence in slide.TimeLine.InteractiveSequences)
            {
                foreach (Effect effect in sequence)
                {
                    if(effect.Shape == shape)
                    {
                        effect.Delete();
                    }
                }
            }

        }
    }
}
