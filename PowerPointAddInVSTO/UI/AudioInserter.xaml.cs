using Microsoft.Office.Interop.PowerPoint;
using PowerPointAddInVSTO.Extensions;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

namespace PowerPointAddInVSTO.UI
{
    /// <summary>
    /// Interaction logic for AudioInserter.xaml
    /// </summary>
    public partial class AudioInserter : Window
    {
        public AudioInserter()
        {
            InitializeComponent();

            var addIn = Globals.ThisAddIn;

            slidetrack.ItemsSource = addIn.GetSlides();
        }

        public string FilePath { get; private set; } = "none";

        private void Browse(object sender, RoutedEventArgs e)
        {
            var addIn = Globals.ThisAddIn;
            var currentSlide = (Slide)slidetrack.CurrentItem;

            Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();
            if (openFileDlg.ShowDialog() == true)
            {
                if (openFileDlg.FileName.EndsWith(".mp3") || openFileDlg.FileName.EndsWith(".wav"))
                {
                    string currentFullAudioName = openFileDlg.SafeFileName;
                    int formatSep = currentFullAudioName.LastIndexOf(".");
                    string currentAudioName = currentFullAudioName.Remove(formatSep);
                    IEnumerable<string> mediaNames = addIn.Application.ActivePresentation.GetMediaNames();
                    if (!mediaNames.Contains(currentAudioName))
                    {
                        FilePath = openFileDlg.FileName;
                        addIn.SetAudio(currentSlide, FilePath);
                        return;
                    }
                    else MessageBox.Show("This audio file has already in your presentation!");
                }
                else MessageBox.Show("Choose [.mp3] or [.wav] type of the file!");
            }
        }

    }
}
