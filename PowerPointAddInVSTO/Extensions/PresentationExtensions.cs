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

        public static float[] GetTimes(this Presentation presentation, string tagName)
        {
            StringBuilder allTimings = new StringBuilder();
            foreach (Slide slide in presentation.Slides)
            {
                if(slide.Tags[tagName].Length > 0)
                {
                    allTimings.Append(slide.Tags[tagName]);
                }
            }
            if (allTimings.Length > 1)
            {
                var newstr = allTimings.ToString().Substring(1);
                float[] timingsOnSlide = Array.ConvertAll(newstr.Split('|'), float.Parse);
                return timingsOnSlide;
            }
            return new float[0];
        }

        public static void ConvertExistTimelineTags(this Presentation presentation)
        {
            foreach (Slide slide in presentation.Slides)
            {
                string oldTimeline = slide.Tags["HST_TIMELINE"];
                if (oldTimeline.Length == 0) continue;
                string newTimeline = oldTimeline.Insert(0, "|");
                slide.Tags.Add("TIMELINE", newTimeline);
            }
        }
    }
}
