using System.Collections.Generic;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddInVSTO.Extensions
{
    public static class SlideExtensions
    {
        public static IEnumerable<Shape> GetAnimatedShapes(this Slide slide)
        {
            foreach(Shape shape in slide.Shapes)
            {
                if(shape.AnimationSettings.Animate == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    yield return shape;
                }
            }
        }

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
    }
}
