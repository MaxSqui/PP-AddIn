using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddInVSTO.Extensions;
using Application = Microsoft.Office.Interop.PowerPoint.Application;

namespace PowerPointAddInVSTO
{
    public partial class ThisAddIn
    {

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
            this.Application.PresentationOpen += Application_PresentationOpen;
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Application.Presentations.Open("C:/Users/maxbe/source/repos/PowerPointAddInVSTO/5197_Graca_JJ.pptx");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        private void Application_PresentationOpen(Presentation Pres)
        {
            Pres.ConvertExistTimelineTags();
        }

    }
}
