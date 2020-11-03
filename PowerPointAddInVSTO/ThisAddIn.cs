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
using System.IO;

namespace PowerPointAddInVSTO
{
    public partial class ThisAddIn
    {

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Application.Presentations.Open("C:/Users/maxbe/source/repos/PowerPointAddInVSTO/5197_Graca_JJ.pptx");
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

            this.Application.PresentationNewSlide += Application_PresentationNewSlide;

            this.Application.PresentationOpen += Application_PresentationOpen;

        }

        private void Application_PresentationOpen(Presentation Pres)
        {
           IEnumerable<string> oldTimings = Pres.GetTags();
            if (oldTimings.Count() > 1)
            {
                MessageBox.Show("You already have timings from old narrator.");
            }
        }

        private void Application_PresentationNewSlide(Slide Sld)
        {
            int p = Sld.Tags.Count;
        }

        public List<Slide> GetSlides()
        {
            IEnumerable<float> s  = Application.ActivePresentation.Slides[1].GetTimings();
            IEnumerable<float> s2 = Application.ActivePresentation.Slides[2].GetTimings();

            if (s != null) Application.ActivePresentation.Slides[1].Tags.Delete("TIMING");
            if (s2 != null) Application.ActivePresentation.Slides[2].Tags.Delete("TIMING");

            Application.ActivePresentation.Slides[1].Tags.Add("TIMING", "|1|2");
            Application.ActivePresentation.Slides[2].Tags.Add("TIMING", "|1|2");
            Application.ActivePresentation.Slides[2].SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;
            Application.ActivePresentation.Slides[1].SlideShowTransition.AdvanceOnTime = MsoTriState.msoTrue;

            List<Slide> slides = new List<Slide>();

            foreach (Slide slide in this.Application.ActivePresentation.Slides)
            {
                slides.Add(slide);
            }
            return slides;
        }

        public IEnumerable<Shape> GetAnimateShapes()
        {
            return Application.ActivePresentation.GetAnimateShapes();
        }

        public void SetAudio(Slide Sld, string path)
        {
            Shape existAudio = Sld.GetAudioShape();
            if (existAudio != null) existAudio.Delete();
            Shape audio = Sld.Shapes.AddMediaObject2(path, MsoTriState.msoTrue, MsoTriState.msoTrue, 750, 500);
            Effect audioEffect = Sld.TimeLine.MainSequence.AddEffect(audio, MsoAnimEffect.msoAnimEffectMediaPlay, MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            audioEffect.MoveTo(1);
            Sld.RemoveAnimationTrigger(audio);
        }

        public MediaBookmark SetBookMark(Shape audio, float durationTime, bool isMin, string bookMarkName)
        {
            const int fromSec = 1000;
            const int fromMin = 60000;
            if (isMin) durationTime = durationTime * fromMin;
            else durationTime = durationTime * fromSec;
            if (durationTime < audio.MediaFormat.Length && durationTime>0)
            {
                MediaBookmark newBookmark = audio.MediaFormat.MediaBookmarks.Add((int)durationTime, bookMarkName);
                return newBookmark;
            }
            return null;

        }

        public void TriggerShapeToBookmark(Slide Sld, Shape shape, Shape audio, MediaBookmark bookmark)
        {
            Sequence objSequence = Sld.TimeLine.InteractiveSequences.Add();
            objSequence.AddTriggerEffect(shape, MsoAnimEffect.msoAnimEffectZoom, MsoAnimTriggerType.msoAnimTriggerOnMediaBookmark, audio, bookmark.Name);
        }


        #endregion
    }
}
