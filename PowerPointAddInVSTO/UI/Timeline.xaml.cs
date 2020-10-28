using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddInVSTO.Extensions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
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
            Effect effect = (Effect)ShapesTimeline.SelectedItem;
            effect.Paragraph = comboBox.SelectedIndex;
        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            Effect currentEffect = (Effect)ShapesTimeline.SelectedItem;
            var shape = currentEffect.Shape;
            TextBox valueBox = (TextBox)sender;
            currentEffect.Timing.TriggerDelayTime = 0;
            float value;
            if (!float.TryParse(valueBox.Text, out value) && value > 0 && value < 65000)
            {
                //TODO: create valication
            }
            else currentEffect.Timing.TriggerDelayTime = value;

        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Effect effect = (Effect)ShapesTimeline.SelectedItem;
            Shape shape = effect.Shape;
            Slide sld = effect.Shape.Parent as Slide;
            string bookmarkName = effect.Index.ToString();
            Shape audio = sld.GetAudioShape();
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
            bool isMin;
            try
            {
                isMin = Convert.ToBoolean(effect.Paragraph);
            }
            catch
            {
                return;
            }

            MediaBookmark newBookmark = addIn.SetBookMark(sld.GetAudioShape(), time, isMin, bookmarkName);
            if (newBookmark == null)
            {
                MessageBox.Show("Input timing is out of the current timing audio");
                return;
            }
            sld.RemoveAnimationTrigger(shape);
            sld.TimeLine.InteractiveSequences
                .Add()
                .AddTriggerEffect(shape, effect.EffectType, MsoAnimTriggerType.msoAnimTriggerOnMediaBookmark, audio, newBookmark.Name);

            ShapesTimeline.CommitEdit();
            ShapesTimeline.CommitEdit();
            ShapesTimeline.Items.Refresh();
        }

    }
}