using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddInVSTO.Extensions;
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
            ShapesTimeline.ItemsSource = addIn.Application.ActivePresentation.GetEffects();
            SlideInfo.ItemsSource = addIn.GetSlides();

        }
        private void SlideInfo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (SlideInfo.SelectedIndex != -1)
            {
                Slide row = (Slide)SlideInfo.SelectedItem;
                ShapesTimeline.ItemsSource = addIn.Application.ActivePresentation.GetEffectsBySlide(row);
                ShapesTimeline.Items.Refresh();
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            ShapesTimeline.ItemsSource = addIn.Application.ActivePresentation.GetEffects();
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
            Effect effect = (Effect)ShapesTimeline.SelectedItem;
            var Shape = effect.Shape;
            Slide sld = effect.Shape.Parent as Slide;
            IEnumerable<Effect> dependentEffects = sld.TimeLine.MainSequence.GetDependentEffects(effect);
            TextBox valueBox = (TextBox)sender;
            effect.Timing.TriggerDelayTime = 0;
            float value;
            if (!float.TryParse(valueBox.Text, out value) && value > 0 && value < 65000)
            {
                //TODO: create valication
            }

            List<float> timings = sld.GetTimingsTag().ToList();
            List<Effect> effects = addIn.Application.ActivePresentation.GetEffectsBySlide(sld).ToList();
            int k = effects.IndexOf(effect);
            float diffrence = value;
            timings[k + 1] = timings[k + 1] + timings[k] - value;
            timings[k] = diffrence;
            //effect.Timing.TriggerDelayTime = value - tags[k];
            sld.Tags.Delete("TIMING");
            sld.Tags.Add("TIMING", sld.ConvertToString(timings));
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Effect effect = (Effect)ShapesTimeline.SelectedItem;
            var effectType = effect.EffectType;
            Shape shape = effect.Shape;
            Slide sld = effect.Shape.Parent as Slide;
            List <float> tags = sld.GetTimingsTag().ToList();
            List<Effect> effects = addIn.Application.ActivePresentation.GetEffects().ToList();
            int k = effects.IndexOf(effect);

            string bookmarkName = effect.Index.ToString();
            Shape audio = sld.GetAudioShape();
            //IEnumerable<Effect> dependentEffects = sld.GetDependentEffects(effect);

            IEnumerable<Effect> dependentEffects = sld.TimeLine.MainSequence.GetDependentEffects(effect);


            if (audio == null)
            {
                AudioInserter audioInserter = new AudioInserter();
                audioInserter.Show();
                MessageBox.Show("This slide do not contain an audio file. Please insert audio in open window");
                return;
            }
            MediaBookmark currentBookmark = audio.MediaFormat.MediaBookmarks.GetBookmark(bookmarkName);
            if (currentBookmark != null) currentBookmark.Delete();

            float time = effect.Timing.TriggerDelayTime;
            tags[k] = time;
            sld.Tags.Delete("HST_TIMELINE");
            sld.Tags.Add("HST_TIMELINE", sld.ConvertToString(tags));
            MediaBookmark newBookmark = addIn.SetBookMark(sld.GetAudioShape(), time, false, bookmarkName);
            if (newBookmark == null)
            {
                MessageBox.Show("Input timing is out of the current timing audio");
                return;
            }
            sld.RemoveAnimationTrigger(shape);
            //Effect f = sld.TimeLine.InteractiveSequences
            //    .Add()
            Effect newEffect = sld.TimeLine.MainSequence.AddEffect(shape, effectType, effect.EffectInformation.BuildByLevelEffect, effect.Timing.TriggerType);
            newEffect.MoveAfter(effect);
            newEffect.Exit = effect.Exit;
            if (shape.Type == MsoShapeType.msoPlaceholder)
            {
                try
                {
                    newEffect.Paragraph = effect.Paragraph;
                }
                catch
                {

                }
            }

            effect.Delete();
            //    .AddTriggerEffect(shape, effectType, MsoAnimTriggerType.msoAnimTriggerOnMediaBookmark, audio, newBookmark.Name);
            foreach (Effect dependentEffect in dependentEffects)
            {
                var newDependentEffect = sld.TimeLine.MainSequence.AddEffect(dependentEffect.Shape, dependentEffect.EffectType, dependentEffect.EffectInformation.BuildByLevelEffect, dependentEffect.Timing.TriggerType);
                newDependentEffect.MoveAfter(dependentEffect);
                newDependentEffect.Exit = dependentEffect.Exit;
                if (dependentEffect.Shape.Type == MsoShapeType.msoPlaceholder)
                {
                    try
                    {
                        newDependentEffect.Paragraph = dependentEffect.Paragraph;
                    }
                    catch
                    {

                    }
                }
                dependentEffect.Delete();
            }
            ShapesTimeline.CommitEdit();
            ShapesTimeline.CommitEdit();
            ShapesTimeline.Items.Refresh();
        }

    }
}