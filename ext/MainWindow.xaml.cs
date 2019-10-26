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
using System.Collections.ObjectModel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
using System.Diagnostics;
using System.Security.Principal;

namespace AAMPCList
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    /// 

    public partial class MainWindow : Window
    {
        int count = 0;
        int alt_count = 6;
        bool edit_mode = false;
        bool move_state = false;
        List<string> tile_list = new List<string>();
        string curr_tile;
    
        public MainWindow()
        {
            InitializeComponent();
            string targetPath = AppDomain.CurrentDomain.BaseDirectory + "\\backup.txt";
            if (File.Exists(targetPath))
            {
                String line;
                StreamReader sr = new StreamReader(targetPath);
                line = sr.ReadLine();
                while (line != null)
                {
                    combo1.Text = line;
                    add_tiles();
                    line = sr.ReadLine();
                }
                sr.Close();
                Console.ReadLine();
            }
            this.MouseLeftButtonDown += new MouseButtonEventHandler(HandleClick);
            this.PreviewKeyDown += new KeyEventHandler(HandleEsc);
        }

        private void HandleEsc(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
                Close();
        }

        private void HandleClick(object sender, MouseButtonEventArgs e)
        {
            if (combo1.Visibility == Visibility.Visible)
            {
                combo1.Visibility = Visibility.Collapsed;
                combo1.IsDropDownOpen = false;
                GrayB.Visibility = Visibility.Collapsed;
            }
        }

        private void Window_closed(object sender, EventArgs e)
        {
            string targetPath = AppDomain.CurrentDomain.BaseDirectory + "\\backup.txt";
            if (File.Exists(targetPath))
            {
                StreamWriter sw = new StreamWriter(targetPath);
                foreach (var list_item in tile_list)
                {
                    sw.WriteLine(list_item);
                }
                sw.Close();
            }
        }

        private void btnBorder_Click1(object sender, RoutedEventArgs e)
        {
            PlexOpt pl1 = new PlexOpt();
            pl1.Show();
        }

        private void btnBorder_Click0(object sender, RoutedEventArgs e)
        {
            if (GridPlex.Visibility == Visibility.Visible)
            {
                Cb1.Visibility = Visibility.Collapsed;
            }
            else
            {
                Cb1.Visibility = Visibility.Visible;
            }

            if (GridWork.Visibility == Visibility.Visible)
            {
                Cb2.Visibility = Visibility.Collapsed;
            }
            else
            {
                Cb2.Visibility = Visibility.Visible;
            }

            if (GridADP.Visibility == Visibility.Visible)
            {
                Cb3.Visibility = Visibility.Collapsed;
            }
            else
            {
                Cb3.Visibility = Visibility.Visible;
            }

            if (GridSelf.Visibility == Visibility.Visible)
            {
                Cb4.Visibility = Visibility.Collapsed;
            }
            else
            {
                Cb4.Visibility = Visibility.Visible;
            }

            if (GridPLM.Visibility == Visibility.Visible)
            {
                Cb5.Visibility = Visibility.Collapsed;
            }
            else
            {
                Cb5.Visibility = Visibility.Visible;
            }

            if (GridOracle.Visibility == Visibility.Visible)
            {
                Cb6.Visibility = Visibility.Collapsed;
            }
            else
            {
                Cb6.Visibility = Visibility.Visible;
            }

            if (GridOffice.Visibility == Visibility.Visible)
            {
                Cb7.Visibility = Visibility.Collapsed;
            }
            else
            {
                Cb7.Visibility = Visibility.Visible;
            }

            if (GridInstall.Visibility == Visibility.Visible)
            {
                Cb8.Visibility = Visibility.Collapsed;
            }
            else
            {
                Cb8.Visibility = Visibility.Visible;
            }

            if (GridVisual.Visibility == Visibility.Visible)
            {
                Cb9.Visibility = Visibility.Collapsed;
            }
            else
            {
                Cb9.Visibility = Visibility.Visible;
            }

            if (GridCalculator.Visibility == Visibility.Visible)
            {
                Cb10.Visibility = Visibility.Collapsed;
            }
            else
            {
                Cb10.Visibility = Visibility.Visible;
            }

            if (GridNotepad.Visibility == Visibility.Visible)
            {
                Cb11.Visibility = Visibility.Collapsed;
            }
            else
            {
                Cb11.Visibility = Visibility.Visible;
            }

            if (GridChrome.Visibility == Visibility.Visible)
            {
                Cb12.Visibility = Visibility.Collapsed;
            }
            else
            {
                Cb12.Visibility = Visibility.Visible;
            }
            if (GridInternet.Visibility == Visibility.Visible)
            {
                Cb13.Visibility = Visibility.Collapsed;
            }
            else
            {
                Cb13.Visibility = Visibility.Visible;
            }
            GrayB.Visibility = Visibility.Visible;
            combo1.Visibility = Visibility.Visible;
            combo1.IsDropDownOpen = true;
        }

        private void btnBorder_Click2(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.myworkday.com/wday/authgwy/aampower/login.htmld");
        }

        private void btnBorder_Click3(object sender, RoutedEventArgs e)
        {
            string targetPath = "\\PLM_Utility\\PLMLaunchMenu.exe";
            if (File.Exists(targetPath))
            {
                ProcessStartInfo startInfo = new ProcessStartInfo("\\PLM_Utility\\PLMLaunchMenu.exe");
                Process p;
                p = Process.Start(startInfo);
            }
            else
            {
                MessageBox.Show("PLM is not installed, please use the installer.", "Error");
            }

        }

        private void btnBorder_Click4(object sender, RoutedEventArgs e)
        {
            SelfServiceOpt ssO = new SelfServiceOpt();
            ssO.Show();
        }

        private void btnBorder_Click5(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://my.adp.com/static/redbox/login.html");
        }

        private void btnBorder_Click6(object sender, RoutedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://r12ebs.aam.com/");
        }

        private void btnBorder_Click7(object sender, RoutedEventArgs e)
        {
            if (this.IsElevated)
            {
                SearchList ss = new SearchList();
                ss.Show();
            }
            else
            {
                MessageBoxResult result = MessageBox.Show("This feature requries admin privileges. Please run this program as admnistrator.", "Warning");
            }
        }


        private void btnBorder_Click8(object sender, RoutedEventArgs e)
        {
            Window1 ss = new Window1();
            ss.Show();
        }

        private void Visual_Click(object sender, RoutedEventArgs e)
        {
            string targetPath = "\\Program Files (x86)\\Microsoft Visual Studio\\2019\\Community\\Common7\\IDE\\devenv.exe";
            if (File.Exists(targetPath))
            {
                ProcessStartInfo startInfo = new ProcessStartInfo(targetPath);
                Process p;
                p = Process.Start(startInfo);
            }
            else
            {
                MessageBox.Show("Visual Studio is not installed, please use the installer.", "Error");
            }
        }

        private void Calc_Click(object sender, RoutedEventArgs e)
        {
            string targetPath = "\\Windows\\System32\\calc.exe";
            if (File.Exists(targetPath))
            {
                ProcessStartInfo startInfo = new ProcessStartInfo(targetPath);
                Process p;
                p = Process.Start(startInfo);
            }
            else
            {
                MessageBox.Show("Calculator is not installed, please use the installer.", "Error");
            }
        }

        private void Note_Click(object sender, RoutedEventArgs e)
        {
            string targetPath = "\\Windows\\System32\\notepad.exe";
            if (File.Exists(targetPath))
            {
                ProcessStartInfo startInfo = new ProcessStartInfo(targetPath);
                Process p;
                p = Process.Start(startInfo);
            }
            else
            {
                MessageBox.Show("Notepad is not installed, please use the installer.", "Error");
            }
        }

        private void Chrome_Click(object sender, RoutedEventArgs e)
        {
            string targetPath = "\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe";
            if (File.Exists(targetPath))
            {
                ProcessStartInfo startInfo = new ProcessStartInfo(targetPath);
                Process p;
                p = Process.Start(startInfo);
            }
            else
            {
                MessageBox.Show("Google Chrome is not installed, please use the installer.", "Error");
            }
        }

        private void Internet_Click(object sender, RoutedEventArgs e)
        {
            string targetPath = "\\Program Files\\internet explorer\\iexplore.exe";
            if (File.Exists(targetPath))
            {
                ProcessStartInfo startInfo = new ProcessStartInfo(targetPath);
                Process p;
                p = Process.Start(startInfo);
            }
            else
            {
                MessageBox.Show("Internet Explorer is not installed, please use the installer.", "Error");
            }
        }

        public bool IsElevated
        {
            get
            {
                return new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
            }
        }

        private void IconClosed(object sender, EventArgs e)
        {
            if (IconSize.Text == "Small")
            {
                Grid1.Height = 421;
                Grid1.Width = 456;
            }
            else if (IconSize.Text == "Medium")
            {
                Grid1.Height = 521;
                Grid1.Width = 656;
            }
            else
            {
                Grid1.Height = 621;
                Grid1.Width = 756;
            }
        }
    }
}
