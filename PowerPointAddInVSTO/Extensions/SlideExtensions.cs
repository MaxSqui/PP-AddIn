using System;
using System.Collections.Generic;
using System.Text;
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

        public static IEnumerable<float> GetTimings(this Slide slide)
        {
            string timingsValueStr = slide.Tags["TIMING"];
            if (timingsValueStr.Length > 0)
            {
                var newstr = timingsValueStr.Substring(1);
                IEnumerable <float> timingsOnSlide = Array.ConvertAll(newstr.Split('|'), float.Parse);
                return timingsOnSlide;
            }
            return null;
        }

        public static IEnumerable<float> GetTags1(this Slide slide)
        {
            string timingsValueStr = slide.Tags["HST_TIMELINE"];
            if (timingsValueStr.Length > 0)
            {
                IEnumerable<float> timingsOnSlide = Array.ConvertAll(timingsValueStr.Split('|'), float.Parse);
                return timingsOnSlide;
            }
            return null;
        }

        //TODO change location
        public static string ConvertToString(this Slide slide, List<float> tag)
        {
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < tag.Count; i++)
            {
                sb.Append(tag[i]);
                sb.Append("|");
                if (i == tag.Count-1)
                {
                    sb.Length -= 1;
                }
            }
            sb.Insert(0, "|");
            return sb.ToString();
        }

        public static string ConvertToString1(this Slide slide, List<float> tag)
        {
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < tag.Count; i++)
            {
                sb.Append(tag[i]);
                sb.Append("|");
                if (i == tag.Count - 1)
                {
                    sb.Length -= 1;
                }
            }
            return sb.ToString();
        }

        public static float GetCurrentTiming(this Slide slide, List<float> timeline, float effectTimeline, int effectPosition)
        {
            float timelineSum = 0;
            float currentTiming = 0;
            if(timeline != null)
            {
                if (timeline.Count >= effectPosition)
                {
                    for (int i = 0; i < effectPosition; i++)
                    {
                        timelineSum += timeline[i];
                    }
                    //TODO Set new timeline value inside tag
                }
                else
                {
                    for (int i = 0; i < timeline.Count; i++)
                    {
                        timelineSum += timeline[i];
                        //TODO Set new timeline value inside tag
                    }
                }

            }
            currentTiming = effectTimeline - timelineSum; 
            return currentTiming;
        }
    }
}
