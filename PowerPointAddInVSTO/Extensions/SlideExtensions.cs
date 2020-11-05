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

        public static float[] GetTimes(this Slide slide, string tagName)
        {
            string timingsValueStr = slide.Tags[tagName];
            if (timingsValueStr.Length > 0)
            {
                var newstr = timingsValueStr.Substring(1);
                float[] timingsOnSlide = Array.ConvertAll(newstr.Split('|'), float.Parse);
                return timingsOnSlide;
            }
            return null;
        }


        //TODO change location
        public static string ConvertTimesToString(this Slide slide, List<float> timings)
        {
            StringBuilder sb = new StringBuilder();

            for (int i = 0; i < timings.Count; i++)
            {
                sb.Append(timings[i]);
                sb.Append("|");
                if (i == timings.Count-1)
                {
                    sb.Length -= 1;
                }
            }
            sb.Insert(0, "|");
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
                }
                else
                {
                    for (int i = 0; i < timeline.Count; i++)
                    {
                        timelineSum += timeline[i];
                    }
                }

            }
            currentTiming = effectTimeline - timelineSum; 
            return currentTiming;
        }

        public static IEnumerable<Effect> GetEffects(this Slide slide)
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
}
