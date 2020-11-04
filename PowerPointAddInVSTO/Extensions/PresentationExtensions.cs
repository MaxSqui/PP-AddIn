using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddInVSTO.ViewModel;
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

        public static IEnumerable<Effect> GetEffects(this Presentation presentation)
        {
            foreach (Slide slide in presentation.Slides)
            {
                foreach (Effect effect in slide.TimeLine.MainSequence)
                {
                    if (effect.Shape.Type != Microsoft.Office.Core.MsoShapeType.msoMedia && effect.Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerOnPageClick)
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
                if (effect.Shape.Type != Microsoft.Office.Core.MsoShapeType.msoMedia && effect.Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerOnPageClick)
                {
                    yield return effect;
                }
            }
        }

        public static IEnumerable<string> GetTags(this Presentation presentation)
        {
            foreach (Slide slide in presentation.Slides)
            {
                yield return slide.Tags["HST_TIMELINE"];
            }
        }

        public static float[] GetTimings(this Presentation presentation)
        {
            StringBuilder allTimings = new StringBuilder();
            foreach (Slide slide in presentation.Slides)
            {
                allTimings.Append(slide.Tags["HST_TIMELINE"]);
                allTimings.Append("|");
            }
            if (allTimings.Length > 0)
            {
                var newstr = allTimings.ToString().Substring(1, allTimings.Length -5);
                float[] timingsOnSlide = Array.ConvertAll(newstr.Split('|'), float.Parse);
                return timingsOnSlide;
            }
            return null;
        }

        public static float[] GetTimingsBySlide(this Presentation presentation, Slide slide)
        {
            StringBuilder allTimings = new StringBuilder();
            allTimings.Append(slide.Tags["HST_TIMELINE"]);
            allTimings.Append("|");

            if (allTimings.Length > 1)
            {
                var newstr = allTimings.ToString().Substring(0, allTimings.Length - 1);
                float[] timingsOnSlide = Array.ConvertAll(newstr.Split('|'), float.Parse);
                return timingsOnSlide;
            }
            return null;
        }
    }
}
