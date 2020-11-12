using System;
using System.Collections.Generic;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointAddInVSTO.Extensions
{
    public static class SlideExtensions
    {
        public static Shape GetAudioShape(this Slide slide)
        {
            foreach (Shape shape in slide.Shapes)
            {
                if (shape.Type == MsoShapeType.msoMedia)
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
        public static double[] GetTimes(this Slide slide, string tagName)
        {
            string timingsValueStr = slide.Tags[tagName];
            if (timingsValueStr.Length > 0)
            {
                var newstr = timingsValueStr.Substring(1);
                double[] timingsOnSlide = Array.ConvertAll(newstr.Split('|'), double.Parse);
                return timingsOnSlide;
            }
            return new double[0];
        }
        public static double GetCurrentTiming(this Slide slide, List<double> timeline, double effectTimeline, int effectPosition)
        {
            double timelineSum = 0;
            double currentTiming = 0;
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
        public static IEnumerable<Effect> GetMainEffects(this Slide slide)
        {
            foreach (Effect effect in slide.TimeLine.MainSequence)
            {
                if (effect.Shape.Type != Microsoft.Office.Core.MsoShapeType.msoMedia && effect.Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerOnPageClick)
                {
                    yield return effect;
                }
            }
        }
        public static void SetAudio(this Slide Sld, string path)
        {
            Shape existAudio = Sld.GetAudioShape();
            if (existAudio != null) existAudio.Delete();
            Shape audio = Sld.Shapes.AddMediaObject2(path, MsoTriState.msoTrue, MsoTriState.msoTrue, 750, 500);
            Effect audioEffect = Sld.TimeLine.MainSequence.AddEffect(audio, MsoAnimEffect.msoAnimEffectMediaPlay, MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            audioEffect.MoveTo(1);
            Sld.RemoveAnimationTrigger(audio);
        }
    }
}
