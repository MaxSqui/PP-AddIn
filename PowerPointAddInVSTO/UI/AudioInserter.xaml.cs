using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddInVSTO.Extensions;
using PowerPointAddInVSTO.ViewModel;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace PowerPointAddInVSTO.UI
{
    public partial class AudioInserter : Window
    {
        ThisAddIn addIn = Globals.ThisAddIn;
        public AudioInserter()
        {
            InitializeComponent();

            slidetrack.ItemsSource = GetSlidetrackViewModels().ToList();
        }
        private void Browse(object sender, RoutedEventArgs e)
        {
            var currentSlidetrackViewModel = slidetrack.CurrentItem as SlidetrackViewModel;
            var currentSlide = addIn.Application.ActivePresentation.GetSlideByNumber(currentSlidetrackViewModel.SlideNumber);

            Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();
            if (openFileDlg.ShowDialog() == true)
            {
                if (openFileDlg.FileName.EndsWith(".mp3") || openFileDlg.FileName.EndsWith(".wav"))
                {
                    string currentFullAudioName = openFileDlg.SafeFileName;
                    int formatSep = currentFullAudioName.LastIndexOf(".");
                    string currentAudioName = currentFullAudioName.Remove(formatSep);
                    IEnumerable<string> mediaNames = addIn.Application.ActivePresentation.GetMediaNames();
                    if (!IsAlreadyInput(mediaNames.ToList(), currentAudioName, currentSlidetrackViewModel.SlideNumber))
                    {
                        currentSlide.Name = openFileDlg.FileName;
                        currentSlidetrackViewModel.AudioPath = openFileDlg.FileName;
                        slidetrack.CommitEdit();
                        slidetrack.CommitEdit();
                        slidetrack.Items.Refresh();
                        currentSlide.SetAudio(currentSlidetrackViewModel.AudioPath);
                        return;
                    }
                    else MessageBox.Show("This audio file has already in your presentation!");
                }
                else MessageBox.Show("Choose [.mp3] or [.wav] type of the file!");
            }
        }

        public bool IsAlreadyInput(List<string> mediaNames, string audioName, int currentSlideNumber)
        {
            for (int i = 0; i < mediaNames.Count(); i++)
            {
                if (mediaNames[i] == audioName && i+1 != currentSlideNumber) return true; 
            }
            return false;
        }

        public IEnumerable<SlidetrackViewModel> GetSlidetrackViewModels()
        {
            IEnumerable<Slide> slides = addIn.Application.ActivePresentation.GetSlides();
            foreach (Slide slide in slides)
            {
                var currentSlidetrackViewModel = new SlidetrackViewModel
                {
                    SlideNumber = slide.SlideNumber,
                    AudioPath = slide.Name
                };
                yield return currentSlidetrackViewModel;
            }
        }
    }
}
