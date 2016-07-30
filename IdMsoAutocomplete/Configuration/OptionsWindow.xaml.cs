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
using IdMsoAutocomplete.Configuration;

namespace IdMsoAutocomplete
{
    /// <summary>
    /// Interaction logic for OptionsWindow.xaml
    /// </summary>
    public partial class OptionsWindow : Window
    {
        public OptionsWindow(Options options)
        {
            InitializeComponent();

            DataContext = options;
            OfficeVersion.ItemsSource = Enum.GetValues(typeof(OfficeVersion)).Cast<OfficeVersion>();
            OfficeApplication.ItemsSource = Enum.GetValues(typeof(OfficeApplication)).Cast<OfficeApplication>();
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
    }
}
