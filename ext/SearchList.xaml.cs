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
using System.IO;
using System.Diagnostics;

namespace AAMPCList
{
    /// <summary>
    /// Interaction logic for SearchList.xaml
    /// </summary>
    public partial class SearchList : Window
    {
        public SearchList()
        {
            InitializeComponent();
        }

        private void txtNameToSearch_TextChanged(object sender,
    TextChangedEventArgs e)
        {
            CollectionViewSource.GetDefaultView(lstEmpData.ItemsSource).Refresh();
        }


        List<string> lstEmployee =
        new List<string>();

        List<string> lstEmployee1 = new List<string>();

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            lstEmployee.Add("Plex");
            lstEmployee.Add("PLM");
            lstEmployee.Add("Visual Studio");
            lstEmployee.Add("Google Chrome");
            lstEmployee.Add("Java");
            lstEmployee.Add("Screen Share");
            lstEmployee.Add("GoToMeeting");
            lstEmployee.Add("Zoom");
            lstEmpData.ItemsSource = lstEmployee;

            lstEmployee1.Add("Plex Browser Plugin");
            lstEmployee1.Add("Plex IE Settings");
            lstEmployee1.Add("Plex Websocket Plugin");
            lstEmpData1.ItemsSource = lstEmployee1;
            lstEmpData1.SelectedItem = -1;
            CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(lstEmpData.ItemsSource);
            view.Filter = UserFilter;

        }

        private bool UserFilter(object obj)
        {
            if (string.IsNullOrEmpty(txtNameToSearch.Text))
            {
                return true;
            }
            else
            {
                return (obj.ToString().IndexOf(txtNameToSearch.Text, StringComparison.OrdinalIgnoreCase) >= 0);
            }
        }

        private void OnSelected(object sender, RoutedEventArgs e)
        {
            if (lstEmpData.SelectedItem != null)
            {
                string temp = lstEmpData.SelectedItem.ToString();
                if (temp == "Plex")
                {
                    lstEmpData1.Visibility = Visibility.Visible;
                    B1.Visibility = Visibility.Collapsed;
                }
                else if (temp == "PLM")
                {
                    lstEmpData1.Visibility = Visibility.Collapsed;
                    B1.Visibility = Visibility.Visible;
                    B2.Visibility = Visibility.Collapsed;
                }
                else if (temp == "Visual Studio")
                {
                    lstEmpData1.Visibility = Visibility.Collapsed;
                    B1.Visibility = Visibility.Visible;
                    B2.Visibility = Visibility.Collapsed;
                }
                else if (temp == "Google Chrome")
                {
                    lstEmpData1.Visibility = Visibility.Collapsed;
                    B1.Visibility = Visibility.Visible;
                    B2.Visibility = Visibility.Collapsed;
                }
                else if (temp == "Java")
                {
                    lstEmpData1.Visibility = Visibility.Collapsed;
                    B1.Visibility = Visibility.Visible;
                    B2.Visibility = Visibility.Collapsed;
                }
                else if (temp == "Screen Share")
                {
                    lstEmpData1.Visibility = Visibility.Collapsed;
                    B1.Visibility = Visibility.Visible;
                    B2.Visibility = Visibility.Collapsed;
                }
                else if (temp == "GoToMeeting")
                {
                    lstEmpData1.Visibility = Visibility.Collapsed;
                    B1.Visibility = Visibility.Visible;
                    B2.Visibility = Visibility.Collapsed;
                }
                else if (temp == "Zoom")
                {
                    lstEmpData1.Visibility = Visibility.Collapsed;
                    B1.Visibility = Visibility.Visible;
                    B2.Visibility = Visibility.Collapsed;
                }
            }
        }

        private void OnSelected1(object sender, RoutedEventArgs e)
        {
            B2.Visibility = Visibility.Visible;
        }

        private void B1Click(object sender, RoutedEventArgs e)
        {
            string temp = lstEmpData.SelectedItem.ToString();
            if (temp == "PLM")
            {
                string sourcePath = AppDomain.CurrentDomain.BaseDirectory + "\\PLM_Utility\\PLMLaunchMenu.exe";
                string targetPath = "\\PLM_Utility\\PLMLaunchMenu.exe";
                if (File.Exists(targetPath))
                {
                    MessageBox.Show("PLM is already installed.");
                }
                else
                {
                    Directory.CreateDirectory("\\PLM_Utility");
                    System.IO.File.Copy(sourcePath, targetPath);
                    MessageBox.Show("PLM was succeessfully installed");
                }
            }
            else if (temp == "Visual Studio")
            {
                string sourcePath = AppDomain.CurrentDomain.BaseDirectory + "\\vs_community.exe";
                if (File.Exists(sourcePath))
                {
                    ProcessStartInfo startInfo = new ProcessStartInfo(sourcePath);
                    Process p;
                    p = Process.Start(startInfo);
                }
                else
                {
                    MessageBox.Show("Error: Missing Visual Studio Installation Files");
                }
            }
            else if (temp == "Google Chrome")
            {
                string sourcePath = AppDomain.CurrentDomain.BaseDirectory + "\\ChromeSetup.exe";
                if (File.Exists(sourcePath))
                {
                    ProcessStartInfo startInfo = new ProcessStartInfo(sourcePath);
                    Process P;
                    P = Process.Start(startInfo);
                }
                else
                {
                    MessageBox.Show("Error: Missing Google Chrome Installation Files");
                }
            }
            else if (temp == "Java")
            {
                string sourcePath = AppDomain.CurrentDomain.BaseDirectory + "\\JavaSetup8u221.exe";
                if (File.Exists(sourcePath))
                {
                    ProcessStartInfo startInfo = new ProcessStartInfo(sourcePath);
                    Process P;
                    P = Process.Start(startInfo);
                }
                else
                {
                    MessageBox.Show("Error: Missing Java Installation Files");
                }
            }
            else if (temp == "Screen Share")
            {
                string sourcePath = AppDomain.CurrentDomain.BaseDirectory + "\\ScreenleapInst.exe";
                if (File.Exists(sourcePath))
                {
                    ProcessStartInfo startInfo = new ProcessStartInfo(sourcePath);
                    Process P;
                    P = Process.Start(startInfo);
                }
                else
                {
                    MessageBox.Show("Error: Missing Screen Share Installation Files");
                }
            }
            else if (temp == "GoToMeeting")
            {
                string sourcePath = AppDomain.CurrentDomain.BaseDirectory + "\\GoToMeeting Installer.exe";
                if (File.Exists(sourcePath))
                {
                    ProcessStartInfo startInfo = new ProcessStartInfo(sourcePath);
                    Process P;
                    P = Process.Start(startInfo);
                }
                else
                {
                    MessageBox.Show("Error: Missing GoToMeeting Installation Files");
                }
            }
            else if (temp == "Zoom")
            {
                string sourcePath = AppDomain.CurrentDomain.BaseDirectory + "\\ZoomInstaller.exe";
                if (File.Exists(sourcePath))
                {
                    ProcessStartInfo startInfo = new ProcessStartInfo(sourcePath);
                    Process P;
                    P = Process.Start(startInfo);
                }
                else
                {
                    MessageBox.Show("Error: Missing Zoom Installation Files");
                }
            }

        }

        private void B2Click(object sender, RoutedEventArgs e)
        {
            if (lstEmpData1.SelectedItem != null)
            {
                string temp = lstEmpData1.SelectedItem.ToString();
                if (temp == "Plex Browser Plugin")
                {
                    ProcessStartInfo startInfo = new ProcessStartInfo(AppDomain.CurrentDomain.BaseDirectory + "\\BrowserPlugin.msi");
                    Process p;
                    p = Process.Start(startInfo);
                }
                else if (temp == "Plex IE Settings")
                {
                    ProcessStartInfo startInfo = new ProcessStartInfo(AppDomain.CurrentDomain.BaseDirectory + "\\Plex_Manufacturing_Cloud_x64_IE_Settings.msi");
                    Process p;
                    p = Process.Start(startInfo);
                }
                else
                {
                    ProcessStartInfo startInfo = new ProcessStartInfo(AppDomain.CurrentDomain.BaseDirectory + "\\Plex_Websocket_Browser_Plugin_x64.msi");
                    Process p;
                    p = Process.Start(startInfo);
                }
            }
        }
    }
}
