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
using System.Windows.Shapes;

namespace AAMPCList
{
    /// <summary>
    /// Interaction logic for SelfServiceOpt.xaml
    /// </summary>
    public partial class SelfServiceOpt : Window
    {
        public SelfServiceOpt()
        {
            InitializeComponent();
            this.PreviewKeyDown += new KeyEventHandler(HandleEsc);
        }

        private void HandleEsc(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }

        private void SelfClick2(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://aamprod.service-now.com");
            Close();
        }

        private void SelfClick1(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://aamprod.service-now.com/self_service");
            Close();
        }
    }
}
