using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddInVSTO.Extensions
{
    public static class PresentationExtensions
    {
        public static IEnumerable<string> GetMediaNames(this Presentation presentation)
        {
            foreach (Slide slide in presentation.Slides)
            {
                foreach (Shape shape in slide.Shapes)
                {
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoMedia)
                    {
                        yield return shape.Name;
                    }
                }
            }
        }

        public static IEnumerable<Shape> GetAnimateShapes(this Presentation presentation)
        {
            foreach (Slide slide in presentation.Slides)
            {
                foreach (Shape shape in slide.Shapes)
                {
                    if (shape.AnimationSettings.Animate == Microsoft.Office.Core.MsoTriState.msoTrue &&
                        shape.Type != Microsoft.Office.Core.MsoShapeType.msoMedia)
                    {
                        yield return shape;
                    }
                }
            }
        }

        public static IEnumerable<Shape> GetShapeRanges(this Presentation presentation)
        {
            foreach (Slide slide in presentation.Slides)
            {
                foreach (Shape shapeRange in slide.Background)
                {
                    yield return shapeRange;
                }
            }
        }

        public static IEnumerable<Effect> GetEffects(this Presentation presentation)
        {
            foreach (Slide slide in presentation.Slides)
            {
                foreach (Effect effect in slide.TimeLine.MainSequence)
                {
                    if (effect.Shape.Type != Microsoft.Office.Core.MsoShapeType.msoMedia)
                    {
                        yield return effect;
                    }
                }
            }
        }

        public static IEnumerable<Effect> GetEffectsBySlide(this Presentation presentation, Slide slide)
        {
            foreach (Effect effect in slide.TimeLine.MainSequence)
            {
                if (effect.Shape.Type != Microsoft.Office.Core.MsoShapeType.msoMedia)
                {
                    yield return effect;
                }
            }
        }

        public static IEnumerable<Microsoft.Office.Core.CustomXMLPart> GetCustomData(this Presentation presentation)
        {
            foreach(Microsoft.Office.Core.CustomXMLPart data in presentation.CustomerData)
            {
                yield return data;
            }
        }
    }
}
