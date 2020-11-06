using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddInVSTO.Extensions;
using PowerPointAddInVSTO.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointAddInVSTO.UI
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class Timeline : Window
    {
        ThisAddIn addIn = Globals.ThisAddIn;
        public Timeline()
        {
            InitializeComponent();
            addIn.Application.ActivePresentation.SetDefaultTimings();
            var effects =addIn.Application.ActivePresentation.GetEffects().ToArray();
            List<EffectViewModel> effectViewModels = new List<EffectViewModel>();
            float[] timeline = null;
            if (addIn.Application.ActivePresentation.GetTimes("TIMELINE").Length > 0)
            {
                timeline = addIn.Application.ActivePresentation.GetTimes("TIMELINE").ToArray();
            }
            for (int i = 0; i < effects.Length; i++)
            {
                var sequence = effects[i].Parent as Sequence;
                var timelineEntity = sequence.Parent as TimeLine;
                var slide = timelineEntity.Parent as Slide;
                if(i >= timeline.Length)
                {
                    EffectViewModel ef = new EffectViewModel
                    {
                        Id = effects[i].Index,
                        DisplayName = effects[i].DisplayName,
                        Slide = slide,
                        SlideNumber = slide.SlideNumber,
                        LastSlideNumber = slide.SlideNumber,
                        Type = effects[i].Shape.Type,
                        EffectTimeline = 0,
                        LastEffectTimeline = 0
                    };
                effectViewModels.Add(ef);
                }
                else if (i> 0)
                {
                    EffectViewModel ef = new EffectViewModel
                    {
                        Id = effects[i].Index,
                        DisplayName = effects[i].DisplayName,
                        Slide = slide,
                        SlideNumber = slide.SlideNumber,
                        LastSlideNumber = (((effects[i - 1].Parent as Sequence).Parent as TimeLine).Parent as Slide).SlideNumber,
                        Type = effects[i].Shape.Type,
                        EffectTimeline = timeline[i],
                        LastEffectTimeline = timeline[i-1]
                    };
                    effectViewModels.Add(ef);
                }
                else
                {
                    EffectViewModel ef = new EffectViewModel
                    {
                        Id = effects[i].Index,
                        DisplayName = effects[i].DisplayName,
                        Slide = slide,
                        SlideNumber = slide.SlideNumber,
                        LastSlideNumber = slide.SlideNumber,
                        Type = effects[i].Shape.Type,
                        EffectTimeline = timeline[i],
                        LastEffectTimeline = 0
                    };
                    effectViewModels.Add(ef);
                }

            }

            
            ShapesTimeline.ItemsSource = effectViewModels;
            SlideInfo.ItemsSource = addIn.GetSlides();

        }
        private void SlideInfo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SlideInfo.SelectedIndex != -1)
            {
                Slide row = (Slide)SlideInfo.SelectedItem;
                var effects = row.GetMainEffects().ToArray();
                List<EffectViewModel> effectViewModels = new List<EffectViewModel>();
                var timeline = row.GetTimes("TIMELINE");
                if (timeline != null) 
                {
                    timeline.ToArray();
                    for (int i = 0; i < effects.Length; i++)
                    {
                        var sequence = effects[i].Parent as Sequence;
                        var timelineEntity = sequence.Parent as TimeLine;
                        var slide = timelineEntity.Parent as Slide;
                        if (i >= timeline.Length)
                        {
                            EffectViewModel ef = new EffectViewModel
                            {
                                Id = effects[i].Index,
                                DisplayName = effects[i].DisplayName,
                                Slide = slide,
                                SlideNumber = slide.SlideNumber,
                                LastSlideNumber = slide.SlideNumber,
                                Type = effects[i].Shape.Type,
                                EffectTimeline = 0,
                                LastEffectTimeline = 0
                            };
                            effectViewModels.Add(ef);
                        }
                        else if(i>0)
                        {
                            EffectViewModel ef = new EffectViewModel
                            {
                                Id = effects[i].Index,
                                DisplayName = effects[i].DisplayName,
                                Slide = slide,
                                SlideNumber = slide.SlideNumber,
                                LastSlideNumber = (((effects[i - 1].Parent as Sequence).Parent as TimeLine).Parent as Slide).SlideNumber,
                                Type = effects[i].Shape.Type,
                                EffectTimeline = timeline[i],
                                LastEffectTimeline = timeline[i - 1]
                            };
                            effectViewModels.Add(ef);
                        }
                        else
                        {
                            EffectViewModel ef = new EffectViewModel
                            {
                                Id = effects[i].Index,
                                DisplayName = effects[i].DisplayName,
                                Slide = slide,
                                SlideNumber = slide.SlideNumber,
                                LastSlideNumber = slide.SlideNumber,
                                Type = effects[i].Shape.Type,
                                EffectTimeline = timeline[i],
                                LastEffectTimeline = 0
                            };
                            effectViewModels.Add(ef);
                        }

                    }
                } 

                ShapesTimeline.ItemsSource = effectViewModels;
                ShapesTimeline.Items.Refresh();
            }
        }

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            var effects = addIn.Application.ActivePresentation.GetEffects().ToArray();
            List<EffectViewModel> effectViewModels = new List<EffectViewModel>();
            float[] timeline = null;
            if (addIn.Application.ActivePresentation.GetTimes("TIMELINE").Length > 0)
            {
                timeline = addIn.Application.ActivePresentation.GetTimes("TIMELINE").ToArray();
            }
            for (int i = 0; i < effects.Length; i++)
            {
                var sequence = effects[i].Parent as Sequence;
                var timelineEntity = sequence.Parent as TimeLine;
                var slide = timelineEntity.Parent as Slide;
                if (i >= timeline.Length)
                {
                    EffectViewModel ef = new EffectViewModel
                    {
                        Id = effects[i].Index,
                        DisplayName = effects[i].DisplayName,
                        Slide = slide,
                        SlideNumber = slide.SlideNumber,
                        LastSlideNumber = slide.SlideNumber,
                        Type = effects[i].Shape.Type,
                        EffectTimeline = 0,
                        LastEffectTimeline = 0
                    };
                    effectViewModels.Add(ef);
                }
                else if (i > 0)
                {
                    EffectViewModel ef = new EffectViewModel
                    {
                        Id = effects[i].Index,
                        DisplayName = effects[i].DisplayName,
                        Slide = slide,
                        SlideNumber = slide.SlideNumber,
                        LastSlideNumber = (((effects[i - 1].Parent as Sequence).Parent as TimeLine).Parent as Slide).SlideNumber,
                        Type = effects[i].Shape.Type,
                        EffectTimeline = timeline[i],
                        LastEffectTimeline = timeline[i - 1]
                    };
                    effectViewModels.Add(ef);
                }
                else
                {
                    EffectViewModel ef = new EffectViewModel
                    {
                        Id = effects[i].Index,
                        DisplayName = effects[i].DisplayName,
                        Slide = slide,
                        SlideNumber = slide.SlideNumber,
                        LastSlideNumber = slide.SlideNumber,
                        Type = effects[i].Shape.Type,
                        EffectTimeline = timeline[i],
                        LastEffectTimeline = 0
                    };
                    effectViewModels.Add(ef);
                }

            }


            ShapesTimeline.ItemsSource = effectViewModels;
            ShapesTimeline.Items.Refresh();

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
            Slide sld = row.Slide;
            List<Effect> effects = sld.GetMainEffects().ToList();
            Effect currentEffect = row.FindEffectById(row.Id, effects);
            List<EffectViewModel> effectViewModels = ShapesTimeline.ItemsSource as List<EffectViewModel>;
            List<float> timings = sld.GetTimes("TIMING").ToList();
            List<float> timeline = sld.GetTimes("TIMELINE").ToList();
            //TODO change logic
            int effectIndex = effects.IndexOf(currentEffect);
            float currentTiming = sld.GetCurrentTiming(timings, row.EffectTimeline, effectIndex);
            if (effectIndex != effects.Count() - 1) timings[effectIndex + 1] = timings[effectIndex + 1] + timings[effectIndex] - currentTiming;
            //TODO validation currvalue < 0
            //if (currentTiming < 0)
            timings[effectIndex] = currentTiming;
            timeline[effectIndex] = row.EffectTimeline;

            sld.Tags.Delete("TIMING");
            sld.Tags.Delete("TIMELINE");
            sld.Tags.Add("TIMING", sld.ConvertTimesToString(timings));
            sld.Tags.Add("TIMELINE", sld.ConvertTimesToString(timeline));
        }
    }
}