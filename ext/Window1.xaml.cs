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
using System.Windows.Media.Animation;
using System.IO;
using System.Diagnostics;
using System.Security.Principal;

namespace AAMPCList
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        public Window1()
        {
            InitializeComponent();
        }

        private void word_click(object sender, RoutedEventArgs e)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo("\\Program Files (x86)\\Microsoft Office\\root\\Office16\\WINWORD.EXE");
            Process p;
            p = Process.Start(startInfo);
            this.Close();
        }

        private void excel_click(object sender, RoutedEventArgs e)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo("\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.EXE");
            Process p;
            p = Process.Start(startInfo);
            this.Close();
        }

        private void powerpoint_click(object sender, RoutedEventArgs e)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo("\\Program Files (x86)\\Microsoft Office\\root\\Office16\\POWERPNT.EXE");
            Process p;
            p = Process.Start(startInfo);
            this.Close();
        }

        private void outlook_click(object sender, RoutedEventArgs e)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo("\\Program Files (x86)\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE");
            Process p;
            p = Process.Start(startInfo);
            this.Close();
        }

        private void onenote_click(object sender, RoutedEventArgs e)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo("\\Program Files (x86)\\Microsoft Office\\root\\Office16\\ONENOTE.EXE");
            Process p;
            p = Process.Start(startInfo);
            this.Close();
        }

        private void skype_click(object sender, RoutedEventArgs e)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo("\\Program Files (x86)\\Microsoft Office\\root\\Office16\\lync.exe");
            Process p;
            p = Process.Start(startInfo);
            this.Close();
        }

    }
}
