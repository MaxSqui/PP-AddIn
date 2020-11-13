using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddInVSTO.Extensions;
using PowerPointAddInVSTO.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace PowerPointAddInVSTO.UI
{
    public partial class Timeline : Window
    {
        public const string TIMING = "TIMING";
        public const string TIMELINE = "TIMELINE";

        ThisAddIn addIn = Globals.ThisAddIn;
        public Timeline()
        {
            InitializeComponent();
            ShapesTimeline.ItemsSource = InitializeEffects().ToList();
            SlideInfo.ItemsSource = InitializeSlideInfos().ToList();
        }
        private void SlideInfo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SlideInfo.SelectedIndex != -1)
            {
                ShapesTimeline.ItemsSource = InitializeEffectsBySlide().ToList();
                ShapesTimeline.Items.Refresh();
            }
        }
        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            ShapesTimeline.ItemsSource = InitializeEffects().ToList();
            SlideInfo.SelectedIndex = -1;
            ShapesTimeline.Items.Refresh();
            SlideInfo.Items.Refresh();

        }
        private void ShapesTimeline_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }   

        private IEnumerable<EffectViewModel> InitializeEffects()
        {
            addIn.Application.ActivePresentation.SetDefaultTimings();
            var effects = addIn.Application.ActivePresentation.GetEffects().ToArray();
            double[] timeline = addIn.Application.ActivePresentation.GetTimes(TIMELINE).ToArray();
            TimeSpan[] timeSpans = addIn.Application.ActivePresentation.GetTimeSpanTimelines(timeline).ToArray();
            for (int effectIndex = 0; effectIndex < effects.Length; effectIndex++)
            {
                var sequence = effects[effectIndex].Parent as Sequence;
                var timelineEntity = sequence.Parent as TimeLine;
                var slide = timelineEntity.Parent as Slide;

                if (effectIndex > 0)
                {
                    EffectViewModel сurrentEffectViewModel = new EffectViewModel
                    {
                        Id = effects[effectIndex].Index,
                        DisplayName = effects[effectIndex].DisplayName,
                        Slide = slide,
                        SlideNumber = slide.SlideNumber,
                        LastSlideNumber = (((effects[effectIndex - 1].Parent as Sequence).Parent as TimeLine).Parent as Slide).SlideNumber,
                        Type = effects[effectIndex].Shape.Type,
                        EffectTimeline = timeSpans[effectIndex],
                        LastEffectTimeline = timeSpans[effectIndex - 1]
                    };
                    yield return сurrentEffectViewModel;
                }
                else
                {
                    EffectViewModel currentEffectViewModel = new EffectViewModel
                    {
                        Id = effects[effectIndex].Index,
                        DisplayName = effects[effectIndex].DisplayName,
                        Slide = slide,
                        SlideNumber = slide.SlideNumber,
                        LastSlideNumber = slide.SlideNumber,
                        Type = effects[effectIndex].Shape.Type,
                        EffectTimeline = timeSpans[effectIndex],
                        LastEffectTimeline = new TimeSpan(0, 0, 0, 0, 0)
                    };
                    yield return currentEffectViewModel;
                }

            }
        }
        private IEnumerable<EffectViewModel> InitializeEffectsBySlide()
        {
            SlideInfoViewModel row = (SlideInfoViewModel)SlideInfo.SelectedItem;
            var effects = row.Slide.GetMainEffects().ToArray();
            var timeline = row.Slide.GetTimes(TIMELINE).ToArray();
            TimeSpan[] timeSpans = addIn.Application.ActivePresentation.GetTimeSpanTimelines(timeline).ToArray();

            for (int effectIndex = 0; effectIndex < effects.Length; effectIndex++)
            {
                var sequence = effects[effectIndex].Parent as Sequence;
                var timelineEntity = sequence.Parent as TimeLine;
                var slide = timelineEntity.Parent as Slide;
                if (effectIndex > 0)
                {
                    EffectViewModel currentEffectViewModel = new EffectViewModel
                    {
                        Id = effects[effectIndex].Index,
                        DisplayName = effects[effectIndex].DisplayName,
                        Slide = slide,
                        SlideNumber = slide.SlideNumber,
                        LastSlideNumber = (((effects[effectIndex - 1].Parent as Sequence).Parent as TimeLine).Parent as Slide).SlideNumber,
                        Type = effects[effectIndex].Shape.Type,
                        EffectTimeline = timeSpans[effectIndex],
                        LastEffectTimeline = timeSpans[effectIndex - 1]
                    };
                    yield return currentEffectViewModel;
                }
                else
                {
                    EffectViewModel currentEffectViewModel = new EffectViewModel
                    {
                        Id = effects[effectIndex].Index,
                        DisplayName = effects[effectIndex].DisplayName,
                        Slide = slide,
                        Effect = effects[effectIndex],
                        SlideNumber = slide.SlideNumber,
                        LastSlideNumber = slide.SlideNumber,
                        Type = effects[effectIndex].Shape.Type,
                        EffectTimeline = timeSpans[effectIndex],
                        LastEffectTimeline = new TimeSpan(0, 0, 0, 0, 0)
                    };
                    yield return currentEffectViewModel;
                }

            }
        }
        private IEnumerable<SlideInfoViewModel> InitializeSlideInfos()
        {
            IEnumerable<Slide> slides = addIn.Application.ActivePresentation.GetSlides();
            foreach (Slide slide in slides)
            {
                var currentSlideInfoViewModel = new SlideInfoViewModel
                {
                    Number = slide.SlideNumber,
                    Clicks = slide.GetMainEffects().Count(),
                    SlideTime = slide.SlideShowTransition.AdvanceTime,
                    Slide = slide
                };
                yield return currentSlideInfoViewModel;
            }
        }

        private void ShapesTimeline_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            EffectViewModel row = (EffectViewModel)ShapesTimeline.SelectedItem;
            Slide sld = row.Slide;
            List<Effect> effects = sld.GetMainEffects().ToList();
            List<double> timings = sld.GetTimes(TIMING).ToList();
            List<double> timeline = sld.GetTimes(TIMELINE).ToList();

            int effectIndex = effects.IndexOf(row.Effect);
            double currentTiming = sld.GetCurrentTiming(timings, row.EffectTimeline.TotalSeconds, effectIndex);
            if (effectIndex != effects.Count() - 1) timings[effectIndex + 1] = timings[effectIndex + 1] + timings[effectIndex] - currentTiming;

            timings[effectIndex] = currentTiming;
            timeline[effectIndex] = (float)row.EffectTimeline.TotalSeconds;

            sld.Tags.Add(TIMING, timings.ConvertTimesToString());
            sld.Tags.Add(TIMELINE, timeline.ConvertTimesToString());
        }

    }
}