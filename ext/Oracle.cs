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
        private void Ora_Del(object sender, RoutedEventArgs e)
        {
            int col = Grid.GetColumn(GridOracle);
            int row = Grid.GetRow(GridOracle);
            int total = (row * 4) + col;
            for (int i = total + 2; i < count; i += 2)
            {
                move_item(i);
            }
            Grid.SetColumn(GridOracle, 3);
            Grid.SetRow(GridOracle, 0);
            GridOracle.Visibility = Visibility.Collapsed;
            OraDel.Visibility = Visibility.Collapsed;
            OraMove.Visibility = Visibility.Collapsed;
            OraGreen.Visibility = Visibility.Collapsed;
            OraRed.Visibility = Visibility.Collapsed;
            Oracle.Background = Brushes.White;
            int countX = (count - 2) % 8;
            int countY = (count - 2) / 8 * 2;
            Grid.SetColumn(AddB, countX);
            Grid.SetRow(AddB, countY);
            tile_list.Remove("Oracle");
            if (count < alt_count && (alt_count - count) > 4)
            {
                alt_count -= 8;
            }
            count -= 2;
            combo1.SelectedIndex = -1;
        }

        private void Ora_Move(object sender, RoutedEventArgs e)
        {
            if (move_state == false)
            {
                OraRed.Visibility = Visibility.Visible;
                OraDel.Visibility = Visibility.Collapsed;
                OraMove.Visibility = Visibility.Collapsed;
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
                if (GridInternet.Visibility == Visibility.Visible)
                {
                    InterGreen.Visibility = Visibility.Visible;
                    InterDel.Visibility = Visibility.Collapsed;
                    InterMove.Visibility = Visibility.Collapsed;
                }
                move_state = true;
                curr_tile = "Oracle";
            }
            else
            {
                int col = Grid.GetColumn(GridOracle);
                int row = Grid.GetRow(GridOracle);

                AddB.Visibility = Visibility.Visible;

                if (curr_tile == "Plex")
                {
                    Grid.SetColumn(GridOracle, Grid.GetColumn(GridPlex));
                    Grid.SetRow(GridOracle, Grid.GetRow(GridPlex));

                    Grid.SetColumn(GridPlex, col);
                    Grid.SetRow(GridPlex, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Oracle");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "PLM")
                {
                    Grid.SetColumn(GridOracle, Grid.GetColumn(GridPLM));
                    Grid.SetRow(GridOracle, Grid.GetRow(GridPLM));

                    Grid.SetColumn(GridPLM, col);
                    Grid.SetRow(GridPLM, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Oracle");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "ADP")
                {
                    Grid.SetColumn(GridOracle, Grid.GetColumn(GridADP));
                    Grid.SetRow(GridOracle, Grid.GetRow(GridADP));

                    Grid.SetColumn(GridADP, col);
                    Grid.SetRow(GridADP, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Oracle");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "Workday")
                {
                    Grid.SetColumn(GridOracle, Grid.GetColumn(GridWork));
                    Grid.SetRow(GridOracle, Grid.GetRow(GridWork));

                    Grid.SetColumn(GridWork, col);
                    Grid.SetRow(GridWork, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Oracle");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "Self Service")
                {
                    Grid.SetColumn(GridOracle, Grid.GetColumn(GridSelf));
                    Grid.SetRow(GridOracle, Grid.GetRow(GridSelf));

                    Grid.SetColumn(GridSelf, col);
                    Grid.SetRow(GridSelf, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Oracle");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "Microsoft Office")
                {
                    Grid.SetColumn(GridOracle, Grid.GetColumn(GridOffice));
                    Grid.SetRow(GridOracle, Grid.GetRow(GridOffice));

                    Grid.SetColumn(GridOffice, col);
                    Grid.SetRow(GridOffice, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Oracle");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "Installer")
                {
                    Grid.SetColumn(GridOracle, Grid.GetColumn(GridInstall));
                    Grid.SetRow(GridOracle, Grid.GetRow(GridInstall));

                    Grid.SetColumn(GridInstall, col);
                    Grid.SetRow(GridInstall, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Oracle");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "Visual Studio")
                {
                    Grid.SetColumn(GridOracle, Grid.GetColumn(GridVisual));
                    Grid.SetRow(GridOracle, Grid.GetRow(GridVisual));

                    Grid.SetColumn(GridVisual, col);
                    Grid.SetRow(GridVisual, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Oracle");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "Calculator")
                {
                    Grid.SetColumn(GridOracle, Grid.GetColumn(GridCalculator));
                    Grid.SetRow(GridOracle, Grid.GetRow(GridCalculator));

                    Grid.SetColumn(GridCalculator, col);
                    Grid.SetRow(GridCalculator, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Oracle");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "Notepad")
                {
                    Grid.SetColumn(GridOracle, Grid.GetColumn(GridNotepad));
                    Grid.SetRow(GridOracle, Grid.GetRow(GridNotepad));

                    Grid.SetColumn(GridNotepad, col);
                    Grid.SetRow(GridNotepad, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Oracle");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "Chrome")
                {
                    Grid.SetColumn(GridOracle, Grid.GetColumn(GridChrome));
                    Grid.SetRow(GridOracle, Grid.GetRow(GridChrome));

                    Grid.SetColumn(GridChrome, col);
                    Grid.SetRow(GridChrome, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Oracle");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                else if (curr_tile == "Internet Explorer")
                {
                    Grid.SetColumn(GridOracle, Grid.GetColumn(GridInternet));
                    Grid.SetRow(GridOracle, Grid.GetRow(GridInternet));

                    Grid.SetColumn(GridInternet, col);
                    Grid.SetRow(GridInternet, row);

                    int in1 = tile_list.IndexOf(curr_tile);
                    int in2 = tile_list.IndexOf("Oracle");
                    string tmp = tile_list[in1];
                    tile_list[in1] = tile_list[in2];
                    tile_list[in2] = tmp;
                }
                move_state = false;
                OraGreen.Visibility = Visibility.Collapsed;
                OraDel.Visibility = Visibility.Visible;
                OraMove.Visibility = Visibility.Visible;
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
                if (GridInternet.Visibility == Visibility.Visible)
                {
                    InterGreen.Visibility = Visibility.Collapsed;
                    InterRed.Visibility = Visibility.Collapsed;
                    InterDel.Visibility = Visibility.Visible;
                    InterMove.Visibility = Visibility.Visible;
                }
            }
        }
    }
}