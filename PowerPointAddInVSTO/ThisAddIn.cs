using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using PowerPointAddInVSTO.Extensions;
using PowerPointAddInVSTO.UI;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using System.Collections.ObjectModel;

namespace PowerPointAddInVSTO
{
    public partial class ThisAddIn
    {

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);

            this.Application.PresentationNewSlide += Application_PresentationNewSlide1;

        }

        public List<Slide> GetSlides()
        {
            List<Slide> slides = new List<Slide>();

            foreach (Slide slide in this.Application.ActivePresentation.Slides)
            {
                slides.Add(slide);
            }
            return slides;
        }
        private void Application_PresentationNewSlide1(Slide Sld)
        {
            var a = new AudioInserter();
            a.Show();
            Shape textBox = Sld.Shapes.AddTextbox(
               Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 50);
            textBox.TextFrame.TextRange.InsertAfter("This text was added by using code.");

            //string path = "C:/Users/maxbe/Downloads/Lil_Uzi_Vert-Baby_Pluto.mp3";
            ////add audio
            //Shape audio = Sld.Shapes.AddMediaObject2(path, MsoTriState.msoTrue);

            ////add settings to the audio (automatic play)
            //Sld.TimeLine.MainSequence.AddEffect(audio, MsoAnimEffect.msoAnimEffectMediaPlay, MsoAnimateByLevel.msoAnimateLevelNone ,MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            //audio.AnimationSettings.PlaySettings.PlayOnEntry = MsoTriState.msoTrue;
            //audio.AnimationSettings.Animate = MsoTriState.msoTrue;
            //audio.AnimationSettings.PlaySettings.LoopUntilStopped = MsoTriState.msoTrue;
            //audio.AnimationSettings.PlaySettings.HideWhileNotPlaying = MsoTriState.msoTrue;

            //add bookmark = duration-value (ms) & name
            //Sequence audioSequence = Sld.TimeLine.InteractiveSequences.Add();
            //audio.MediaFormat.MediaBookmarks.Add(5000, "yeet");

            //creating new shape and bind with exist bookmark
            //Shape rectangle = Sld.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, 200, 160, 100, 50);
            //Sequence objSequence = Sld.TimeLine.InteractiveSequences.Add();
            //objSequence.AddTriggerEffect(rectangle, MsoAnimEffect.msoAnimEffectZoom, MsoAnimTriggerType.msoAnimTriggerOnMediaBookmark, audio, "yeet");
        }

        //separate methods (WIP)

        public void SetAudio(Slide Sld, string path)
        {
            Shape audio = Sld.Shapes.AddMediaObject2(path, MsoTriState.msoTrue);

            Sld.TimeLine.MainSequence.AddEffect(audio, MsoAnimEffect.msoAnimEffectMediaPlay, MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            audio.AnimationSettings.PlaySettings.PlayOnEntry = MsoTriState.msoTrue;
            audio.AnimationSettings.Animate = MsoTriState.msoTrue;
            audio.AnimationSettings.PlaySettings.LoopUntilStopped = MsoTriState.msoTrue;
            audio.AnimationSettings.PlaySettings.HideWhileNotPlaying = MsoTriState.msoTrue;
        }

        private void SetBookMark(Shape audio, double durationTime, bool isSec, string bookMarkName)
        {
            const int fromSec = 1000;
            const int fromMin = 60000;
            if (isSec) durationTime = durationTime * fromSec;
            else durationTime = durationTime * fromMin;

            audio.MediaFormat.MediaBookmarks.Add((int)durationTime, bookMarkName);
        }

        private void TriggerShapeToBookmark(Slide Sld, Shape shape, Shape audio, MediaBookmark bookmark)
        {
            Sequence objSequence = Sld.TimeLine.InteractiveSequences.Add();
            objSequence.AddTriggerEffect(shape, MsoAnimEffect.msoAnimEffectZoom, MsoAnimTriggerType.msoAnimTriggerOnMediaBookmark, audio, bookmark.Name);
        }


        #endregion
    }
}
