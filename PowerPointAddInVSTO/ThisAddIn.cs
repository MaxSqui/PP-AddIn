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

namespace PowerPointAddInVSTO
{
    public partial class ThisAddIn
    {
        void Application_PresentationNewSlide(PowerPoint.Slide Sld, float X, float Y)
        {
            int a = Sld.TimeLine.InteractiveSequences.Count;
            PowerPoint.Shape textBox = Sld.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 50);
            textBox.TextFrame.TextRange.InsertAfter("This text was added by using code.");
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            var myUserControl1 = new UserControl();
            var myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl1, "My Task Pane");
            myCustomTaskPane.Visible = true;
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
            this.Application.AfterDragDropOnSlide += Application_PresentationNewSlide;

            this.Application.PresentationNewSlide += Application_PresentationNewSlide1;

            this.Application.AfterShapeSizeChange += Application_AfterShapeSizeChange;

            this.Application.WindowSelectionChange += Application_WindowSelectionChange;

            this.Application.WindowDeactivate += Application_WindowDeactivate; ;

        }

        private void Application_WindowDeactivate(Presentation Pres, DocumentWindow Wn)
        {

        }

        private void Application_WindowSelectionChange(PowerPoint.Selection Sel)
        {
        }

        private void Application_AfterShapeSizeChange(PowerPoint.Shape shp)
        {

        }

        private void Application_PresentationNewSlide1(Slide Sld)
        {
            Shape textBox = Sld.Shapes.AddTextbox(
               Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 50);
            textBox.TextFrame.TextRange.InsertAfter("This text was added by using code.");

            string path = "C:/Users/maxbe/Downloads/Lil_Uzi_Vert-Baby_Pluto.mp3";
            //add audio
            Shape audio = Sld.Shapes.AddMediaObject2(path, MsoTriState.msoTrue);

            //add settings to the audio 
            //TODO: find the way to set AfterPrevious to play automatically & cross slides
            audio.AnimationSettings.PlaySettings.LoopUntilStopped = MsoTriState.msoTrue;
            audio.AnimationSettings.PlaySettings.HideWhileNotPlaying = MsoTriState.msoTrue;
            audio.AnimationSettings.Animate = MsoTriState.msoTrue;

            //add bookmark = duration-value (ms) & name
            Sequence audioSequence = Sld.TimeLine.InteractiveSequences.Add();
            audio.AnimationSettings.PlaySettings.PlayOnEntry = MsoTriState.msoTrue;
            audio.MediaFormat.MediaBookmarks.Add(5000, "yeet");

            //creating new shape and bind with exist bookmark
            Shape rectangle = Sld.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangle, 200, 160, 100, 50);
            Sequence objSequence = Sld.TimeLine.InteractiveSequences.Add(1);
            objSequence.AddTriggerEffect(rectangle, MsoAnimEffect.msoAnimEffectZoom, MsoAnimTriggerType.msoAnimTriggerOnMediaBookmark, audio, "yeet");
        }

        //separate methods (WIP)

        private void SetAudio(Slide Sld, string path)
        {
            Shape audio = Sld.Shapes.AddMediaObject2(path, MsoTriState.msoTrue);
        }

        private void SetBookMark(Shape audio, double durationTime, bool isSec, string bookMarkName)
        {
            const int fromSec = 1000;
            const int fromMin = 60000;
            if (isSec) durationTime = durationTime * fromSec;
            else durationTime = durationTime * fromMin;

            audio.MediaFormat.MediaBookmarks.Add((int)durationTime, bookMarkName);
        }


        #endregion
    }
}
