using PowerPointAddInVSTO.Extensions;
using System;
using System.Collections.Generic;
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
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class Timeline : Window
    {
        public Timeline()
        {
            var addIn = Globals.ThisAddIn;
            InitializeComponent();
            timeline.ItemsSource = addIn.Application.ActivePresentation.GetEffects();
        }

        private void Apply(object sender, RoutedEventArgs e) { }
    }
}
