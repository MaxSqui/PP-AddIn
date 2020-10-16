using Microsoft.Office.Interop.PowerPoint;
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

        public string FilePath { get; private set; }

        private void Browse(object sender, RoutedEventArgs e)
        {
            var addIn = Globals.ThisAddIn;
            var currentSlide = (Slide)slidetrack.Items.CurrentItem;

            Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();
            if (openFileDlg.ShowDialog() == true)
            {
                this.FilePath = openFileDlg.FileName;
                addIn.SetAudio(currentSlide, FilePath);
            }
        }

    }
}
