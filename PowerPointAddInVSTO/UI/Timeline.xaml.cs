using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddInVSTO.Extensions;
using PowerPointAddInVSTO.ViewModel;
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
            SlideInfo.ItemsSource = addIn.Application.ActivePresentation.GetSlides().ToList();
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
        private void SetTiming_Click(object sender, RoutedEventArgs e)
        {
            ShapesTimeline.CommitEdit();
            ShapesTimeline.CommitEdit();

            EffectViewModel row = (EffectViewModel)ShapesTimeline.SelectedItem;
            if (!row.IsSec) row.EffectTimeline = row.EffectTimeline * 60;
            Slide sld = row.Slide;
            List<Effect> effects = sld.GetMainEffects().ToList();
            Effect currentEffect = row.FindEffectById(row.Id, effects);
            List<float> timings = sld.GetTimes(TIMING).ToList();
            List<float> timeline = sld.GetTimes(TIMELINE).ToList();

            int effectIndex = effects.IndexOf(currentEffect);
            float currentTiming = sld.GetCurrentTiming(timings, row.EffectTimeline, effectIndex);
            if (effectIndex != effects.Count() - 1) timings[effectIndex + 1] = timings[effectIndex + 1] + timings[effectIndex] - currentTiming;

            timings[effectIndex] = currentTiming;
            timeline[effectIndex] = row.EffectTimeline;

            sld.Tags.Delete(TIMING);
            sld.Tags.Delete(TIMELINE);
            sld.Tags.Add(TIMING, timings.ConvertTimesToString());
            sld.Tags.Add(TIMELINE, timeline.ConvertTimesToString());
        }

        private IEnumerable<EffectViewModel> InitializeEffects()
        {
            addIn.Application.ActivePresentation.SetDefaultTimings();
            var effects = addIn.Application.ActivePresentation.GetEffects().ToArray();
            float[] timeline = addIn.Application.ActivePresentation.GetTimes(TIMELINE).ToArray();
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
                        EffectTimeline = timeline[effectIndex],
                        LastEffectTimeline = timeline[effectIndex - 1]
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
                        EffectTimeline = timeline[effectIndex],
                        LastEffectTimeline = 0
                    };
                    yield return currentEffectViewModel;
                }

            }
        }
        private IEnumerable<EffectViewModel> InitializeEffectsBySlide()
        {
            Slide row = (Slide)SlideInfo.SelectedItem;
            var effects = row.GetMainEffects().ToArray();
            var timeline = row.GetTimes(TIMELINE).ToArray();
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
                        EffectTimeline = timeline[effectIndex],
                        LastEffectTimeline = timeline[effectIndex - 1]
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
                        SlideNumber = slide.SlideNumber,
                        LastSlideNumber = slide.SlideNumber,
                        Type = effects[effectIndex].Shape.Type,
                        EffectTimeline = timeline[effectIndex],
                        LastEffectTimeline = 0
                    };
                    yield return currentEffectViewModel;
                }

            }
        }
    }
}