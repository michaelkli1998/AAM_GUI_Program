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
        private void Combo1_DropDownClosed(object sender, EventArgs e)
        {
            if (combo1.SelectedIndex == -1)
            {
                combo1.Visibility = Visibility.Collapsed;
                GrayB.Visibility = Visibility.Collapsed;
                return;
            }
            int countX = 0;
            int countY = 0;
            combo1.Visibility = Visibility.Collapsed;
            GrayB.Visibility = Visibility.Collapsed;
            if (count > 6)
            {
                countX = count % 8;
                countY = count / 8 * 2;
            }
            else
            {
                countX = count;
                countY = 0;
            }

            if (alt_count < 6)
            {
                alt_count = 6;
            }

            if (combo1.SelectedValue.ToString() == "Cb1")
            {
                if (GridPlex.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("Plex is already added.", "Error:");
                    combo1.SelectedIndex = -1;
                }
                else
                {
                    GridPlex.Visibility = Visibility.Visible;
                    if (edit_mode == true)
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
                    Grid.SetColumn(GridPlex, countX);
                    Grid.SetRow(GridPlex, countY);
                    if (count < 6)
                    {
                        Grid.SetColumn(AddB, countX + 2);
                    }
                    else if (count == alt_count)
                    {
                        Grid.SetColumn(AddB, 0);
                        Grid.SetRow(AddB, countY + 2);
                        alt_count += 8;
                    }
                    else
                    {
                        Grid.SetColumn(AddB, countX + 2);
                        Grid.SetRow(AddB, countY);
                    }
                    count += 2;
                    tile_list.Add("Plex");
                    combo1.SelectedIndex = -1;
                }
            }
            else if (combo1.SelectedValue.ToString() == "Cb2")
            {
                if (GridWork.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("Workday is already added.", "Error:");
                    combo1.SelectedIndex = -1;
                }
                else
                {
                    GridWork.Visibility = Visibility.Visible;
                    if (edit_mode == true)
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
                    Grid.SetColumn(GridWork, countX);
                    Grid.SetRow(GridWork, countY);
                    if (count < 6)
                    {
                        Grid.SetColumn(AddB, countX + 2);
                    }
                    else if (count == alt_count)
                    {
                        Grid.SetColumn(AddB, 0);
                        Grid.SetRow(AddB, countY + 2);
                        alt_count += 8;
                    }
                    else
                    {
                        Grid.SetColumn(AddB, countX + 2);
                        Grid.SetRow(AddB, countY);
                    }
                    count += 2;
                    tile_list.Add("Workday");
                    combo1.SelectedIndex = -1;
                }
            }
            else if (combo1.SelectedValue.ToString() == "Cb3")
            {
                if (GridADP.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("ADP is already added.", "Error:");
                    combo1.SelectedIndex = -1;
                }
                else
                {
                    GridADP.Visibility = Visibility.Visible;
                    if (edit_mode == true)
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
                    Grid.SetColumn(GridADP, countX);
                    Grid.SetRow(GridADP, countY);
                    if (count < 6)
                    {
                        Grid.SetColumn(AddB, countX + 2);
                    }
                    else if (count == alt_count)
                    {
                        Grid.SetColumn(AddB, 0);
                        Grid.SetRow(AddB, countY + 2);
                        alt_count += 8;
                    }
                    else
                    {
                        Grid.SetColumn(AddB, countX + 2);
                        Grid.SetRow(AddB, countY);
                    }
                    count += 2;
                    tile_list.Add("ADP");
                    combo1.SelectedIndex = -1;
                }
            }
            else if (combo1.SelectedValue.ToString() == "Cb4")
            {
                if (GridSelf.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("Self Service is already added.", "Error:");
                    combo1.SelectedIndex = -1;
                }
                else
                {
                    GridSelf.Visibility = Visibility.Visible;
                    if (edit_mode == true)
                    {
                        SelfDel.Visibility = Visibility.Visible;
                        SelfMove.Visibility = Visibility.Visible;
                        SelfService.Background = Brushes.LightYellow;
                    }
                    else
                    {
                        SelfDel.Visibility = Visibility.Collapsed;
                        SelfMove.Visibility = Visibility.Collapsed;
                        SelfService.Background = Brushes.White;
                    }
                    Grid.SetColumn(GridSelf, countX);
                    Grid.SetRow(GridSelf, countY);
                    if (count < 6)
                    {
                        Grid.SetColumn(AddB, countX + 2);
                    }
                    else if (count == alt_count)
                    {
                        Grid.SetColumn(AddB, 0);
                        Grid.SetRow(AddB, countY + 2);
                        alt_count += 8;
                    }
                    else
                    {
                        Grid.SetColumn(AddB, countX + 2);
                        Grid.SetRow(AddB, countY);
                    }
                    count += 2;
                    tile_list.Add("Self Service");
                    combo1.SelectedIndex = -1;
                }
            }
            else if (combo1.SelectedValue.ToString() == "Cb5")
            {
                if (GridPLM.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("PLM is already added.", "Error:");
                    combo1.SelectedIndex = -1;
                }
                else
                {
                    GridPLM.Visibility = Visibility.Visible;
                    if (edit_mode == true)
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
                    Grid.SetColumn(GridPLM, countX);
                    Grid.SetRow(GridPLM, countY);
                    if (count < 6)
                    {
                        Grid.SetColumn(AddB, countX + 2);
                    }
                    else if (count == alt_count)
                    {
                        Grid.SetColumn(AddB, 0);
                        Grid.SetRow(AddB, countY + 2);
                        alt_count += 8;
                    }
                    else
                    {
                        Grid.SetColumn(AddB, countX + 2);
                        Grid.SetRow(AddB, countY);
                    }
                    count += 2;
                    tile_list.Add("PLM");
                    combo1.SelectedIndex = -1;
                }
            }
            else if (combo1.SelectedValue.ToString() == "Cb6")
            {
                if (GridOracle.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("Oracle is already added.", "Error:");
                    combo1.SelectedIndex = -1;
                }
                else
                {
                    GridOracle.Visibility = Visibility.Visible;
                    if (edit_mode == true)
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
                    Grid.SetColumn(GridOracle, countX);
                    Grid.SetRow(GridOracle, countY);
                    if (count < 6)
                    {
                        Grid.SetColumn(AddB, countX + 2);
                    }
                    else if (count == alt_count)
                    {
                        Grid.SetColumn(AddB, 0);
                        Grid.SetRow(AddB, countY + 2);
                        alt_count += 8;
                    }
                    else
                    {
                        Grid.SetColumn(AddB, countX + 2);
                        Grid.SetRow(AddB, countY);
                    }
                    count += 2;
                    tile_list.Add("Oracle");
                    combo1.SelectedIndex = -1;
                }
            }
            else if (combo1.SelectedValue.ToString() == "Cb7")
            {
                if (GridOffice.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("Microsoft Office is already added.", "Error:");
                    combo1.SelectedIndex = -1;
                }
                else
                {
                    GridOffice.Visibility = Visibility.Visible;
                    if (edit_mode == true)
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
                    Grid.SetColumn(GridOffice, countX);
                    Grid.SetRow(GridOffice, countY);
                    if (count < 6)
                    {
                        Grid.SetColumn(AddB, countX + 2);
                    }
                    else if (count == alt_count)
                    {
                        Grid.SetColumn(AddB, 0);
                        Grid.SetRow(AddB, countY + 2);
                        alt_count += 8;
                    }
                    else
                    {
                        Grid.SetColumn(AddB, countX + 2);
                        Grid.SetRow(AddB, countY);
                    }
                    count += 2;
                    tile_list.Add("Microsoft Office");
                    combo1.SelectedIndex = -1;
                }
            }
            else if (combo1.SelectedValue.ToString() == "Cb8")
            {
                if (GridInstall.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("The installer is already added.", "Error:");
                    combo1.SelectedIndex = -1;
                }
                else
                {
                    GridInstall.Visibility = Visibility.Visible;
                    if (edit_mode == true)
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
                    Grid.SetColumn(GridInstall, countX);
                    Grid.SetRow(GridInstall, countY);
                    if (count < 6)
                    {
                        Grid.SetColumn(AddB, countX + 2);
                    }
                    else if (count == alt_count)
                    {
                        Grid.SetColumn(AddB, 0);
                        Grid.SetRow(AddB, countY + 2);
                        alt_count += 8;
                    }
                    else
                    {
                        Grid.SetColumn(AddB, countX + 2);
                        Grid.SetRow(AddB, countY);
                    }
                    count += 2;
                    tile_list.Add("Installer");
                    combo1.SelectedIndex = -1;
                }
            }
            else if (combo1.SelectedValue.ToString() == "Cb9")
            {
                if (GridVisual.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("Visual Studio is already added.", "Error:");
                    combo1.SelectedIndex = -1;
                }
                else
                {
                    GridVisual.Visibility = Visibility.Visible;
                    if (edit_mode == true)
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
                    Grid.SetColumn(GridVisual, countX);
                    Grid.SetRow(GridVisual, countY);
                    if (count < 6)
                    {
                        Grid.SetColumn(AddB, countX + 2);
                    }
                    else if (count == alt_count)
                    {
                        Grid.SetColumn(AddB, 0);
                        Grid.SetRow(AddB, countY + 2);
                        alt_count += 8;
                    }
                    else
                    {
                        Grid.SetColumn(AddB, countX + 2);
                        Grid.SetRow(AddB, countY);
                    }
                    count += 2;
                    tile_list.Add("Visual Studio");
                    combo1.SelectedIndex = -1;
                }
            }
            else if (combo1.SelectedValue.ToString() == "Cb10")
            {
                if (GridCalculator.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("Calculator is already added.", "Error:");
                    combo1.SelectedIndex = -1;
                }
                else
                {
                    GridCalculator.Visibility = Visibility.Visible;
                    if (edit_mode == true)
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
                    Grid.SetColumn(GridCalculator, countX);
                    Grid.SetRow(GridCalculator, countY);
                    if (count < 6)
                    {
                        Grid.SetColumn(AddB, countX + 2);
                    }
                    else if (count == alt_count)
                    {
                        Grid.SetColumn(AddB, 0);
                        Grid.SetRow(AddB, countY + 2);
                        alt_count += 8;
                    }
                    else
                    {
                        Grid.SetColumn(AddB, countX + 2);
                        Grid.SetRow(AddB, countY);
                    }
                    count += 2;
                    tile_list.Add("Calculator");
                    combo1.SelectedIndex = -1;
                }
            }
            else if (combo1.SelectedValue.ToString() == "Cb11")
            {
                if (GridNotepad.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("Notepad is already added.", "Error:");
                    combo1.SelectedIndex = -1;
                }
                else
                {
                    GridNotepad.Visibility = Visibility.Visible;
                    if (edit_mode == true)
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
                    Grid.SetColumn(GridNotepad, countX);
                    Grid.SetRow(GridNotepad, countY);
                    if (count < 6)
                    {
                        Grid.SetColumn(AddB, countX + 2);
                    }
                    else if (count == alt_count)
                    {
                        Grid.SetColumn(AddB, 0);
                        Grid.SetRow(AddB, countY + 2);
                        alt_count += 8;
                    }
                    else
                    {
                        Grid.SetColumn(AddB, countX + 2);
                        Grid.SetRow(AddB, countY);
                    }
                    count += 2;
                    tile_list.Add("Notepad");
                    combo1.SelectedIndex = -1;
                }
            }
            else if (combo1.SelectedValue.ToString() == "Cb12")
            {
                if (GridChrome.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("Chrome is already added.", "Error:");
                    combo1.SelectedIndex = -1;
                }
                else
                {
                    GridChrome.Visibility = Visibility.Visible;
                    if (edit_mode == true)
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
                    Grid.SetColumn(GridChrome, countX);
                    Grid.SetRow(GridChrome, countY);
                    if (count < 6)
                    {
                        Grid.SetColumn(AddB, countX + 2);
                    }
                    else if (count == alt_count)
                    {
                        Grid.SetColumn(AddB, 0);
                        Grid.SetRow(AddB, countY + 2);
                        alt_count += 8;
                    }
                    else
                    {
                        Grid.SetColumn(AddB, countX + 2);
                        Grid.SetRow(AddB, countY);
                    }
                    count += 2;
                    tile_list.Add("Chrome");
                    combo1.SelectedIndex = -1;
                }
            }
            else if (combo1.SelectedValue.ToString() == "Cb13")
            {
                if (GridInternet.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("Internet Explorer is already added.", "Error:");
                    combo1.SelectedIndex = -1;
                }
                else
                {
                    GridInternet.Visibility = Visibility.Visible;
                    if (edit_mode == true)
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
                    Grid.SetColumn(GridInternet, countX);
                    Grid.SetRow(GridInternet, countY);
                    if (count < 6)
                    {
                        Grid.SetColumn(AddB, countX + 2);
                    }
                    else if (count == alt_count)
                    {
                        Grid.SetColumn(AddB, 0);
                        Grid.SetRow(AddB, countY + 2);
                        alt_count += 8;
                    }
                    else
                    {
                        Grid.SetColumn(AddB, countX + 2);
                        Grid.SetRow(AddB, countY);
                    }
                    count += 2;
                    tile_list.Add("Internet Explorer");
                    combo1.SelectedIndex = -1;
                }
            }
        }
    }
}