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

        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            if (edit_mode == false)
            {
                edit_mode = true;
                Edit.Background = Brushes.Red;
                IconSize.Visibility = Visibility.Collapsed;
                IconText.Visibility = Visibility.Collapsed;
                AddB.Visibility = Visibility.Visible;
            }
            else
            {
                edit_mode = false;
                IconSize.Visibility = Visibility.Collapsed;
                IconText.Visibility = Visibility.Collapsed;
                Edit.Background = Brushes.White;
                AddB.Visibility = Visibility.Visible;
                PlexGreen.Visibility = Visibility.Collapsed;
                PlexRed.Visibility = Visibility.Collapsed;
                WorkRed.Visibility = Visibility.Collapsed;
                WorkGreen.Visibility = Visibility.Collapsed;
                ADPGreen.Visibility = Visibility.Collapsed;
                ADPRed.Visibility = Visibility.Collapsed;
                SelfGreen.Visibility = Visibility.Collapsed;
                SelfRed.Visibility = Visibility.Collapsed;
                PLMGreen.Visibility = Visibility.Collapsed;
                PLMRed.Visibility = Visibility.Collapsed;
                OraGreen.Visibility = Visibility.Collapsed;
                OraRed.Visibility = Visibility.Collapsed;
                OfficeGreen.Visibility = Visibility.Collapsed;
                OfficeRed.Visibility = Visibility.Collapsed;
                InstGreen.Visibility = Visibility.Collapsed;
                InstRed.Visibility = Visibility.Collapsed;
                VisualGreen.Visibility = Visibility.Collapsed;
                VisualRed.Visibility = Visibility.Collapsed;
                CalcGreen.Visibility = Visibility.Collapsed;
                CalcRed.Visibility = Visibility.Collapsed;
                NoteGreen.Visibility = Visibility.Collapsed;
                NoteRed.Visibility = Visibility.Collapsed;
                ChromeGreen.Visibility = Visibility.Collapsed;
                ChromeRed.Visibility = Visibility.Collapsed;
                InterGreen.Visibility = Visibility.Collapsed;
                InterRed.Visibility = Visibility.Collapsed;
                move_state = false;
            }

            if (GridPlex.Visibility == Visibility.Visible && edit_mode == true)
            {
                PlexDel.Visibility = Visibility.Visible;
                PlexMove.Visibility = Visibility.Visible;
                Plex.Background = Brushes.LightYellow;
            }
            else
            {
                PlexDel.Visibility = Visibility.Collapsed;
                PlexMove.Visibility = Visibility.Collapsed;
                Plex.Background = Brushes.White;
            }

            if (GridWork.Visibility == Visibility.Visible && edit_mode == true)
            {
                WorkDel.Visibility = Visibility.Visible;
                WorkMove.Visibility = Visibility.Visible;
                Workday.Background = Brushes.LightYellow;
            }
            else
            {
                WorkDel.Visibility = Visibility.Collapsed;
                WorkMove.Visibility = Visibility.Collapsed;
                Workday.Background = Brushes.White;
            }

            if (GridPLM.Visibility == Visibility.Visible && edit_mode == true)
            {
                PLMDel.Visibility = Visibility.Visible;
                PLMMove.Visibility = Visibility.Visible;
                PLM.Background = Brushes.LightYellow;
            }
            else
            {
                PLMDel.Visibility = Visibility.Collapsed;
                PLMMove.Visibility = Visibility.Collapsed;
                PLM.Background = Brushes.White;
            }

            if (GridADP.Visibility == Visibility.Visible && edit_mode == true)
            {
                ADPDel.Visibility = Visibility.Visible;
                ADPMove.Visibility = Visibility.Visible;
                ADP.Background = Brushes.LightYellow;
            }
            else
            {
                ADPDel.Visibility = Visibility.Collapsed;
                ADPMove.Visibility = Visibility.Collapsed;
                ADP.Background = Brushes.White;
            }

            if (GridSelf.Visibility == Visibility.Visible && edit_mode == true)
            {
                SelfDel.Visibility = Visibility.Visible;
                SelfMove.Visibility = Visibility.Visible;
                SelfService.Background = Brushes.LightYellow;
                SelfText.Visibility = Visibility.Collapsed;
            }
            else
            {
                SelfDel.Visibility = Visibility.Collapsed;
                SelfMove.Visibility = Visibility.Collapsed;
                SelfText.Visibility = Visibility.Visible;
                SelfService.Background = Brushes.White;
            }

            if (GridOracle.Visibility == Visibility.Visible && edit_mode == true)
            {
                OraDel.Visibility = Visibility.Visible;
                OraMove.Visibility = Visibility.Visible;
                Oracle.Background = Brushes.LightYellow;
            }
            else
            {
                OraDel.Visibility = Visibility.Collapsed;
                OraMove.Visibility = Visibility.Collapsed;
                Oracle.Background = Brushes.White;
            }

            if (GridOffice.Visibility == Visibility.Visible && edit_mode == true)
            {
                OfficeDel.Visibility = Visibility.Visible;
                OfficeMove.Visibility = Visibility.Visible;
                Office.Background = Brushes.LightYellow;
            }
            else
            {
                OfficeDel.Visibility = Visibility.Collapsed;
                OfficeMove.Visibility = Visibility.Collapsed;
                Office.Background = Brushes.White;
            }

            if (GridInstall.Visibility == Visibility.Visible && edit_mode == true)
            {
                InstDel.Visibility = Visibility.Visible;
                InstMove.Visibility = Visibility.Visible;
                Install.Background = Brushes.LightYellow;
            }
            else
            {
                InstDel.Visibility = Visibility.Collapsed;
                InstMove.Visibility = Visibility.Collapsed;
                Install.Background = Brushes.White;
            }

            if (GridVisual.Visibility == Visibility.Visible && edit_mode == true)
            {
                VisualDel.Visibility = Visibility.Visible;
                VisualMove.Visibility = Visibility.Visible;
                Visual.Background = Brushes.LightYellow;
            }
            else
            {
                VisualDel.Visibility = Visibility.Collapsed;
                VisualMove.Visibility = Visibility.Collapsed;
                Visual.Background = Brushes.White;
            }

            if (GridCalculator.Visibility == Visibility.Visible && edit_mode == true)
            {
                CalcDel.Visibility = Visibility.Visible;
                CalcMove.Visibility = Visibility.Visible;
                Calculator.Background = Brushes.LightYellow;
            }
            else
            {
                CalcDel.Visibility = Visibility.Collapsed;
                CalcMove.Visibility = Visibility.Collapsed;
                Calculator.Background = Brushes.White;
            }

            if (GridNotepad.Visibility == Visibility.Visible && edit_mode == true)
            {
                NoteDel.Visibility = Visibility.Visible;
                NoteMove.Visibility = Visibility.Visible;
                Notepad.Background = Brushes.LightYellow;
            }
            else
            {
                NoteDel.Visibility = Visibility.Collapsed;
                NoteMove.Visibility = Visibility.Collapsed;
                Notepad.Background = Brushes.White;
            }
            
            if (GridChrome.Visibility == Visibility.Visible && edit_mode == true)
            {
                ChromeDel.Visibility = Visibility.Visible;
                ChromeMove.Visibility = Visibility.Visible;
                Chrome.Background = Brushes.LightYellow;
            }
            else
            {
                ChromeDel.Visibility = Visibility.Collapsed;
                ChromeMove.Visibility = Visibility.Collapsed;
                Chrome.Background = Brushes.White;
            }

            if (GridInternet.Visibility == Visibility.Visible && edit_mode == true)
            {
                InterDel.Visibility = Visibility.Visible;
                InterMove.Visibility = Visibility.Visible;
                Internet.Background = Brushes.LightYellow;
            }
            else
            {
                InterDel.Visibility = Visibility.Collapsed;
                InterMove.Visibility = Visibility.Collapsed;
                Internet.Background = Brushes.White;
            }
        }
    }
}