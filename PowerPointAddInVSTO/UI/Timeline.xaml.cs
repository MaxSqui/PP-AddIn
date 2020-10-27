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
            //bool isSec = bool.Parse(comboBox.SelectedIndex.ToString());
            Effect effect = (Effect)ShapesTimeline.SelectedItem;
            float value = effect.Timing.TriggerDelayTime;
            Slide sld = effect.Shape.Parent as Slide;
            Shape audio = sld.GetAudioShape();
            MsoShapeType t = audio.Type;
            if (audio.MediaFormat == null)
            {

            }
            addIn.SetBookMark(sld.GetAudioShape(), effect.Timing.TriggerDelayTime, true);

        }

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox valueBox = (TextBox)sender;
            float value = float.Parse(valueBox.Text);
            Effect currentEffect = (Effect)ShapesTimeline.SelectedItem;
            currentEffect.Timing.TriggerDelayTime = value;
        }
    }
}