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
            var effects =addIn.Application.ActivePresentation.GetEffects().ToArray();
            List<EffectViewModel> effectViewModels = new List<EffectViewModel>();
            var timings = addIn.Application.ActivePresentation.GetTimings().ToArray();
            for(int i = 0; i < effects.Length; i++)
            {
                var sequence = effects[i].Parent as Sequence;
                var timeline = sequence.Parent as TimeLine;
                var slide = timeline.Parent as Slide;
                if (i >= timings.Length)
                {
                    EffectViewModel ef = new EffectViewModel
                    {
                        DisplayName = effects[i].DisplayName,
                        Slide = slide,
                        SlideNumber = slide.SlideNumber,
                        Type = effects[i].Shape.Type,
                        EffectTimeline = 0
                    };
                    effectViewModels.Add(ef);
                }
                else
                {
                    EffectViewModel ef = new EffectViewModel
                    {
                        DisplayName = effects[i].DisplayName,
                        Slide = slide,
                        SlideNumber = slide.SlideNumber,
                        Type = effects[i].Shape.Type,
                        EffectTimeline = timings[i]

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
                var effects = addIn.Application.ActivePresentation.GetEffectsBySlide(row).ToArray();
                List<EffectViewModel> effectViewModels = new List<EffectViewModel>();
                var timings = addIn.Application.ActivePresentation.GetTimingsBySlide(row);
                if (timings != null) 
                {
                    timings.ToArray();
                    for (int i = 0; i < effects.Length; i++)
                    {
                        var sequence = effects[i].Parent as Sequence;
                        var timeline = sequence.Parent as TimeLine;
                        var slide = timeline.Parent as Slide;
                        if (i >= timings.Length)
                        {
                            EffectViewModel ef = new EffectViewModel
                            {
                                DisplayName = effects[i].DisplayName,
                                Slide = slide,
                                SlideNumber = slide.SlideNumber,
                                Type = effects[i].Shape.Type,
                                EffectTimeline = 0
                            };
                            effectViewModels.Add(ef);
                        }
                        else
                        {
                            EffectViewModel ef = new EffectViewModel
                            {
                                DisplayName = effects[i].DisplayName,
                                Slide = slide,
                                SlideNumber = slide.SlideNumber,
                                Type = effects[i].Shape.Type,
                                EffectTimeline = timings[i]

                            };
                            effectViewModels.Add(ef);
                        }

                    }
                } 

                ShapesTimeline.ItemsSource = effectViewModels;
                ShapesTimeline.Items.Refresh();
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var effects = addIn.Application.ActivePresentation.GetEffects().ToArray();
            List<EffectViewModel> effectViewModels = new List<EffectViewModel>();
            var timings = addIn.Application.ActivePresentation.GetTimings().ToArray();
            for (int i = 0; i < effects.Length; i++)
            {
                var sequence = effects[i].Parent as Sequence;
                var timeline = sequence.Parent as TimeLine;
                var slide = timeline.Parent as Slide;
                if (i >= timings.Length)
                {
                    EffectViewModel ef = new EffectViewModel
                    {
                        DisplayName = effects[i].DisplayName,
                        Slide = slide,
                        SlideNumber = slide.SlideNumber,
                        Type = effects[i].Shape.Type,
                        EffectTimeline = 0
                    };
                    effectViewModels.Add(ef);
                }
                else
                {
                    EffectViewModel ef = new EffectViewModel
                    {
                        DisplayName = effects[i].DisplayName,
                        Slide = slide,
                        SlideNumber = slide.SlideNumber,
                        Type = effects[i].Shape.Type,
                        EffectTimeline = timings[i]

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

        private void ActionCombo_SelectionChanged(object sender, SelectionChangedEventArgs e) 
        {
            ComboBox comboBox = (ComboBox)sender;
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            EffectViewModel effectViewModel = ShapesTimeline.SelectedItem as EffectViewModel;

            Slide sld = effectViewModel.Slide;
            TextBox valueBox = (TextBox)sender;
            float value;
            if (!float.TryParse(valueBox.Text, out value) && value > 0 && value < 65000)
            {
                //TODO: create valication
            }
            List<float> timings = sld.GetTimings().ToList();
            List<EffectViewModel> effectViewModels = ShapesTimeline.ItemsSource as List<EffectViewModel>;

            int effectIndex = effectViewModels.IndexOf(effectViewModel);
            float currvalue = sld.GetCurrentTiming(timings, value, effectIndex);
            if (effectIndex!=effectViewModels.Count-1) timings[effectIndex + 1] = timings[effectIndex + 1] + timings[effectIndex] - value;
            //TODO validation currvalue < 0
            timings[effectIndex] = currvalue;
            sld.Tags.Delete("TIMING");
            sld.Tags.Add("TIMING", sld.ConvertToString(timings));
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            ShapesTimeline.CommitEdit();
            ShapesTimeline.CommitEdit();
            ShapesTimeline.Items.Refresh();
        }

        private void ShapesTimeline_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}