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
            Application.Presentations.Open("C:/Users/maxbe/source/repos/PowerPointAddInVSTO/5197_Graca_JJ.pptx");
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

        public IEnumerable<Shape> GetAnimateShapes()
        {
            return Application.ActivePresentation.GetAnimateShapes();
        }
        private void Application_PresentationNewSlide1(Slide Sld)
        {

        }

        //separate methods (WIP)

        public void SetAudio(Slide Sld, string path)
        {
            Shape existAudio = Sld.GetAudioShape();
            if (existAudio != null) existAudio.Delete();
            Shape audio = Sld.Shapes.AddMediaObject2(path, MsoTriState.msoTrue);
            Effect audioEffect = Sld.TimeLine.MainSequence.AddEffect(audio, MsoAnimEffect.msoAnimEffectMediaPlay, MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            audioEffect.MoveTo(1);
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
