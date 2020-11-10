using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointAddInVSTO.Extensions
{
    public static class PresentationExtensions
    {
        public const string TIMING = "TIMING";
        public const string TIMELINE = "TIMELINE";
        public const string OLD_TIMELINE = "HST_TIMELINE";
        public static IEnumerable<Slide> GetSlides(this Presentation presentation)
        {
            foreach (Slide slide in presentation.Slides)
            {
                yield return slide;
            }
        }
        public static Slide GetSlideByNumber(this Presentation presentation, int slideNumber)
        {
            foreach (Slide slide in presentation.Slides)
            {
                if (slide.SlideNumber == slideNumber)
                    return slide; 
            }
            return null;
        }
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
                string oldTimeline = slide.Tags[OLD_TIMELINE];
                if (oldTimeline.Length == 0) continue;
                string newTimeline = oldTimeline.Insert(0, "|");
                slide.Tags.Add(TIMELINE, newTimeline);
            }
        }

        public static void SetDefaultTimings(this Presentation presentation)
        {
            foreach (Slide slide in presentation.Slides)
            {
                IEnumerable<Effect> effects = slide.GetMainEffects();
                string timingsTag = slide.Tags[TIMING];
                string timelineTag = slide.Tags[TIMELINE];
               
                if (timingsTag.Length == 0)
                {
                    StringBuilder times = new StringBuilder();
                    foreach (Effect effect in effects)
                    {
                        times.Append("|0");
                    }
                    slide.SlideShowTransition.AdvanceOnTime = Microsoft.Office.Core.MsoTriState.msoTrue;
                    slide.Tags.Add(TIMING, times.ToString());
                    slide.Tags.Add(TIMELINE, times.ToString());
                }

                float[] timings = slide.GetTimes(TIMING);

                if (effects.Count() > timings.Length)
                {
                    StringBuilder timingsBuilder = new StringBuilder(timingsTag);
                    StringBuilder timelineBuilder = new StringBuilder(timelineTag);
                    int diffrence = effects.Count() - timings.Length;
                    for (int i = 0; i < diffrence; i++)
                    {
                        timingsBuilder.Append('|');
                        timingsBuilder.Append('0');

                        timelineBuilder.Append('|');
                        timelineBuilder.Append('0');

                    }
                    slide.Tags.Delete(TIMING);
                    slide.Tags.Delete(TIMELINE);
                    slide.Tags.Add(TIMING, timingsBuilder.ToString());
                    slide.Tags.Add(TIMELINE, timelineBuilder.ToString());
                }

                if(effects.Count() < timings.Length)
                {
                    int diffrence =  timings.Length- effects.Count();
                    for (int i = 0; i < diffrence; i++)
                    {
                        int timingSeparator = timingsTag.LastIndexOf('|');
                        timingsTag = timingsTag.Remove(timingSeparator);

                        int timelineSeparator = timelineTag.LastIndexOf('|');
                        timelineTag = timelineTag.Remove(timelineSeparator);

                    }
                    slide.Tags.Delete(TIMING);
                    slide.Tags.Delete(TIMELINE);
                    slide.Tags.Add(TIMING, timingsTag);
                    slide.Tags.Add(TIMELINE, timelineTag);

                }
            }
        }
    }
}
