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
    /// Interaction logic for PlexOpt.xaml
    /// </summary>
    public partial class PlexOpt : Window
    {
        public PlexOpt()
        {
            InitializeComponent();
            this.PreviewKeyDown += new KeyEventHandler(HandleEsc);
        }

        private void HandleEsc(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }

        private void PlexClick2(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://cloud.plex.com");
            Close();
        }

        private void PlexClick1(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.plexus-online.com");
            Close();
        }
    }
}
