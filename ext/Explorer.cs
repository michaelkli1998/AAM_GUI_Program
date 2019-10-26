﻿using System;
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
        private void Inter_Del(object sender, RoutedEventArgs e)
        {
            int col = Grid.GetColumn(GridInternet);
            int row = Grid.GetRow(GridInternet);
            int total = (row * 4) + col;
            for (int i = total + 2; i < count; i += 2)
            {
                move_item(i);
            }
            Grid.SetColumn(GridInternet, 3);
            Grid.SetRow(GridInternet, 0);
            GridInternet.Visibility = Visibility.Collapsed;
            InterDel.Visibility = Visibility.Collapsed;
            InterMove.Visibility = Visibility.Collapsed;
            InterGreen.Visibility = Visibility.Collapsed;
            InterRed.Visibility = Visibility.Collapsed;
            Internet.Background = Brushes.White;
            int countX = (count - 2) % 8;
            int countY = (count - 2) / 8 * 2;
            Grid.SetColumn(AddB, countX);
            Grid.SetRow(AddB, countY);
            tile_list.Remove("Internet Explorer");
            if (count < alt_count && (alt_count - count) > 4)
            {
                alt_count -= 8;
            }
            count -= 2;
            combo1.SelectedIndex = -1;
        }

        private void Inter_Move(object sender, RoutedEventArgs e)
        {
            if (move_state == false)
            {
                InterRed.Visibility = Visibility.Visible;
                InterDel.Visibility = Visibility.Collapsed;
                InterMove.Visibility = Visibility.Collapsed;
                AddB.Visibility = Visibility.Collapsed;
                if (GridPlex.Visibility == Visibility.Visible)
                {
                    PlexGreen.Visibility = Visibility.Visible;
                    PlexDel.Visibility = Visibility.Collapsed;
                    PlexMove.Visibility = Visibility.Collapsed;
                }
                if (GridPLM.Visibility == Visibility.Visible)
                {
                    PLMGreen.Visibility = Visibility.Visible;
                    PLMDel.Visibility = Visibility.Collapsed;
                    PLMMove.Visibility = Visibility.Collapsed;
                }
                if (GridADP.Visibility == Visibility.Visible)
                {
                    ADPGreen.Visibility = Visibility.Visible;
                    ADPDel.Visibility = Visibility.Collapsed;
                    ADPMove.Visibility = Visibility.Collapsed;
                }
                if (GridWork.Visibility == Visibility.Visible)
                {
                    WorkGreen.Visibility = Visibility.Visible;
                    WorkDel.Visibility = Visibility.Collapsed;
                    WorkMove.Visibility = Visibility.Collapsed;
                }
                if (GridSelf.Visibility == Visibility.Visible)
                {
                    SelfGreen.Visibility = Visibility.Visible;
                    SelfDel.Visibility = Visibility.Collapsed;
                    SelfMove.Visibility = Visibility.Collapsed;
                }
                if (GridOracle.Visibility == Visibility.Visible)
                {
                    OraGreen.Visibility = Visibility.Visible;
                    OraDel.Visibility = Visibility.Collapsed;
                    OraMove.Visibility = Visibility.Collapsed;
                }
                if (GridOffice.Visibility == Visibility.Visible)
                {
                    OfficeGreen.Visibility = Visibility.Visible;
                    OfficeDel.Visibility = Visibility.Collapsed;
                    OfficeMove.Visibility = Visibility.Collapsed;
                }
                if (GridInstall.Visibility == Visibility.Visible)
                {
                    InstGreen.Visibility = Visibility.Visible;
                    InstDel.Visibility = Visibility.Collapsed;
                    InstMove.Visibility = Visibility.Collapsed;
                }
                if (GridVisual.Visibility == Visibility.Visible)
                {
                    VisualGreen.Visibility = Visibility.Visible;
                    VisualDel.Visibility = Visibility.Collapsed;
                    VisualMove.Visibility = Visibility.Collapsed;
                }
                if (GridCalculator.Visibility == Visibility.Visible)
                {
                    CalcGreen.Visibility = Visibility.Visible;
                    CalcDel.Visibility = Visibility.Collapsed;
                    CalcMove.Visibility = Visibility.Collapsed;
                }
                if (GridNotepad.Visibility == Visibility.Visible)
                {
                    NoteGreen.Visibility = Visibility.Visible;
                    NoteDel.Visibility = Visibility.Collapsed;
                    NoteMove.Visibility = Visibility.Collapsed;
                }
                if (GridChrome.Visibility == Visibility.Visible)
                {
                    ChromeGreen.Visibility = Visibility.Visible;
                    ChromeDel.Visibility = Visibility.Collapsed;
                    ChromeMove.Visibility = Visibility.Collapsed;
                }
                move_state = true;
                curr_tile = "Internet Explorer";
            }
            else
            {
                int col = Grid.GetColumn(GridInternet);
                int row = Grid.GetRow(GridInternet);

                AddB.Visibility = Visibility.Visible;

                if (curr_tile == "Plex")
                {
                    Grid.SetColumn(GridInternet, Grid.GetColumn(GridPlex));
                    Grid.SetRow(GridInternet, Grid.GetRow(GridPlex));

                    Grid.SetColumn(GridPlex, col);
                    Grid.SetRow(GridPlex, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Internet Explorer");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "PLM")
                {
                    Grid.SetColumn(GridInternet, Grid.GetColumn(GridPLM));
                    Grid.SetRow(GridInternet, Grid.GetRow(GridPLM));

                    Grid.SetColumn(GridPLM, col);
                    Grid.SetRow(GridPLM, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Internet Explorer");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "ADP")
                {
                    Grid.SetColumn(GridInternet, Grid.GetColumn(GridADP));
                    Grid.SetRow(GridInternet, Grid.GetRow(GridADP));

                    Grid.SetColumn(GridADP, col);
                    Grid.SetRow(GridADP, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Internet Explorer");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "Workday")
                {
                    Grid.SetColumn(GridInternet, Grid.GetColumn(GridWork));
                    Grid.SetRow(GridInternet, Grid.GetRow(GridWork));

                    Grid.SetColumn(GridWork, col);
                    Grid.SetRow(GridWork, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Internet Explorer");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "Self Service")
                {
                    Grid.SetColumn(GridInternet, Grid.GetColumn(GridSelf));
                    Grid.SetRow(GridInternet, Grid.GetRow(GridSelf));

                    Grid.SetColumn(GridSelf, col);
                    Grid.SetRow(GridSelf, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Internet Explorer");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "Oracle")
                {
                    Grid.SetColumn(GridInternet, Grid.GetColumn(GridOracle));
                    Grid.SetRow(GridInternet, Grid.GetRow(GridOracle));

                    Grid.SetColumn(GridOracle, col);
                    Grid.SetRow(GridOracle, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Internet Explorer");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "Microsoft Office")
                {
                    Grid.SetColumn(GridInternet, Grid.GetColumn(GridOffice));
                    Grid.SetRow(GridInternet, Grid.GetRow(GridOffice));

                    Grid.SetColumn(GridOffice, col);
                    Grid.SetRow(GridOffice, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Internet Explorer");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "Installer")
                {
                    Grid.SetColumn(GridInternet, Grid.GetColumn(GridInstall));
                    Grid.SetRow(GridInternet, Grid.GetRow(GridInstall));

                    Grid.SetColumn(GridInstall, col);
                    Grid.SetRow(GridInstall, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Internet Explorer");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "Visual Studio")
                {
                    Grid.SetColumn(GridInternet, Grid.GetColumn(GridVisual));
                    Grid.SetRow(GridInternet, Grid.GetRow(GridVisual));

                    Grid.SetColumn(GridVisual, col);
                    Grid.SetRow(GridVisual, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Internet Explorer");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "Calculator")
                {
                    Grid.SetColumn(GridInternet, Grid.GetColumn(GridCalculator));
                    Grid.SetRow(GridInternet, Grid.GetRow(GridCalculator));

                    Grid.SetColumn(GridCalculator, col);
                    Grid.SetRow(GridCalculator, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Internet Explorer");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "Notepad")
                {
                    Grid.SetColumn(GridInternet, Grid.GetColumn(GridNotepad));
                    Grid.SetRow(GridInternet, Grid.GetRow(GridNotepad));

                    Grid.SetColumn(GridNotepad, col);
                    Grid.SetRow(GridNotepad, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Internet Explorer");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "Chrome")
                {
                    Grid.SetColumn(GridInternet, Grid.GetColumn(GridChrome));
                    Grid.SetRow(GridInternet, Grid.GetRow(GridChrome));
 
                    Grid.SetColumn(GridChrome, col);
                    Grid.SetRow(GridChrome, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Internet Explorer");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                move_state = false;
                InterGreen.Visibility = Visibility.Collapsed;
                InterDel.Visibility = Visibility.Visible;
                InterMove.Visibility = Visibility.Visible;
                if (GridWork.Visibility == Visibility.Visible)
                {
                    WorkGreen.Visibility = Visibility.Collapsed;
                    WorkRed.Visibility = Visibility.Collapsed;
                    WorkDel.Visibility = Visibility.Visible;
                    WorkMove.Visibility = Visibility.Visible;
                }
                if (GridPLM.Visibility == Visibility.Visible)
                {
                    PLMGreen.Visibility = Visibility.Collapsed;
                    PLMRed.Visibility = Visibility.Collapsed;
                    PLMDel.Visibility = Visibility.Visible;
                    PLMMove.Visibility = Visibility.Visible;
                }
                if (GridADP.Visibility == Visibility.Visible)
                {
                    ADPGreen.Visibility = Visibility.Collapsed;
                    ADPRed.Visibility = Visibility.Collapsed;
                    ADPDel.Visibility = Visibility.Visible;
                    ADPMove.Visibility = Visibility.Visible;
                }
                if (GridPlex.Visibility == Visibility.Visible)
                {
                    PlexGreen.Visibility = Visibility.Collapsed;
                    PlexRed.Visibility = Visibility.Collapsed;
                    PlexDel.Visibility = Visibility.Visible;
                    PlexMove.Visibility = Visibility.Visible;
                }
                if (GridSelf.Visibility == Visibility.Visible)
                {
                    SelfGreen.Visibility = Visibility.Collapsed;
                    SelfRed.Visibility = Visibility.Collapsed;
                    SelfDel.Visibility = Visibility.Visible;
                    SelfMove.Visibility = Visibility.Visible;
                }
                if (GridOracle.Visibility == Visibility.Visible)
                {
                    OraGreen.Visibility = Visibility.Collapsed;
                    OraRed.Visibility = Visibility.Collapsed;
                    OraDel.Visibility = Visibility.Visible;
                    OraMove.Visibility = Visibility.Visible;
                }
                if (GridOffice.Visibility == Visibility.Visible)
                {
                    OfficeGreen.Visibility = Visibility.Collapsed;
                    OfficeRed.Visibility = Visibility.Collapsed;
                    OfficeDel.Visibility = Visibility.Visible;
                    OfficeMove.Visibility = Visibility.Visible;
                }
                if (GridInstall.Visibility == Visibility.Visible)
                {
                    InstGreen.Visibility = Visibility.Collapsed;
                    InstRed.Visibility = Visibility.Collapsed;
                    InstDel.Visibility = Visibility.Visible;
                    InstMove.Visibility = Visibility.Visible;
                }
                if (GridVisual.Visibility == Visibility.Visible)
                {
                    VisualGreen.Visibility = Visibility.Collapsed;
                    VisualRed.Visibility = Visibility.Collapsed;
                    VisualDel.Visibility = Visibility.Visible;
                    VisualMove.Visibility = Visibility.Visible;
                }
                if (GridCalculator.Visibility == Visibility.Visible)
                {
                    CalcGreen.Visibility = Visibility.Collapsed;
                    CalcRed.Visibility = Visibility.Collapsed;
                    CalcDel.Visibility = Visibility.Visible;
                    CalcMove.Visibility = Visibility.Visible;
                }
                if (GridNotepad.Visibility == Visibility.Visible)
                {
                    NoteGreen.Visibility = Visibility.Collapsed;
                    NoteRed.Visibility = Visibility.Collapsed;
                    NoteDel.Visibility = Visibility.Visible;
                    NoteMove.Visibility = Visibility.Visible;
                }
                if (GridChrome.Visibility == Visibility.Visible)
                {
                    ChromeGreen.Visibility = Visibility.Collapsed;
                    ChromeRed.Visibility = Visibility.Collapsed;
                    ChromeDel.Visibility = Visibility.Visible;
                    ChromeMove.Visibility = Visibility.Visible;
                }
            }
        }
    }
}