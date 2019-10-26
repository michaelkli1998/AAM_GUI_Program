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
        private void add_tiles()
        {
            int countX = 0;
            int countY = 0;
            if (count > 6)
            {
                countX = count % 8;
                countY = count / 8 * 2;
                if (countY % 2 != 0)
                {
                    countY += 1;
                }
            }
            else
            {
                countX = count;
                countY = 0;
            }

            if (combo1.Text == "Plex")
            {
                if (GridPlex.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("Plex is already added.", "Error:");
                }
                else
                {
                    GridPlex.Visibility = Visibility.Visible;
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
                }
            }
            else if (combo1.Text == "Workday")
            {
                if (GridWork.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("Workday is already added.", "Error:");
                }
                else
                {
                    GridWork.Visibility = Visibility.Visible;
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
                }
            }
            else if (combo1.Text == "ADP")
            {
                if (GridADP.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("ADP is already added.", "Error:");
                }
                else
                {
                    GridADP.Visibility = Visibility.Visible;
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
                }
            }
            else if (combo1.Text == "Self Service")
            {
                if (GridSelf.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("Self Service is already added.", "Error:");
                }
                else
                {
                    GridSelf.Visibility = Visibility.Visible;
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
                }
            }
            else if (combo1.Text == "PLM")
            {
                if (GridPLM.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("PLM is already added.", "Error:");
                }
                else
                {
                    GridPLM.Visibility = Visibility.Visible;
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
                }
            }
            else if (combo1.Text == "Oracle")
            {
                if (GridOracle.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("Oracle is already added.", "Error:");
                }
                else
                {
                    GridOracle.Visibility = Visibility.Visible;
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
                }
            }
            else if (combo1.Text == "Microsoft Office")
            {
                if (GridOffice.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("Microsoft Office is already added.", "Error:");
                }
                else
                {
                    GridOffice.Visibility = Visibility.Visible;
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
                }
            }
            else if (combo1.Text == "Installer")
            {
                if (GridInstall.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("The installer is already added.", "Error:");
                }
                else
                {
                    GridInstall.Visibility = Visibility.Visible;
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
                }
            }
            else if (combo1.Text == "Visual Studio")
            {
                if (GridVisual.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("Visual Studio is already added.", "Error:");
                }
                else
                {
                    GridVisual.Visibility = Visibility.Visible;
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
                }
            }
            else if (combo1.Text == "Calculator")
            {
                if (GridCalculator.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("Calculator is already added.", "Error:");
                }
                else
                {
                    GridCalculator.Visibility = Visibility.Visible;
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
                }
            }
            else if (combo1.Text == "Notepad")
            {
                if (GridNotepad.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("Notepad is already added.", "Error:");
                }
                else
                {
                    GridNotepad.Visibility = Visibility.Visible;
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
                }
            }
            else if (combo1.Text == "Chrome")
            {
                if (GridChrome.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("Chrome is already added.", "Error:");
                }
                else
                {
                    GridChrome.Visibility = Visibility.Visible;
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
                }
            }
            else if (combo1.Text == "Internet Explorer")
            {
                if (GridInternet.Visibility == Visibility.Visible)
                {
                    MessageBox.Show("Internet Explorer is already added.", "Error:");
                }
                else
                {
                    GridInternet.Visibility = Visibility.Visible;
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
                }
            }
        }

        private void move_item(int total)
        {
            int col;
            int row;
            int tot;

            col = Grid.GetColumn(GridPlex);
            row = Grid.GetRow(GridPlex);
            tot = (row * 4) + col;
            if (tot == total)
            {
                tot -= 2;
                int countX = tot % 8;
                int countY = tot / 8 * 2;
                Grid.SetColumn(GridPlex, countX);
                Grid.SetRow(GridPlex, countY);
                return;
            }

            col = Grid.GetColumn(GridWork);
            row = Grid.GetRow(GridWork);
            tot = (row * 4) + col;
            if (tot == total)
            {
                tot -= 2;
                int countX = tot % 8;
                int countY = tot / 8 * 2;
                Grid.SetColumn(GridWork, countX);
                Grid.SetRow(GridWork, countY);
                return;
            }

            col = Grid.GetColumn(GridADP);
            row = Grid.GetRow(GridADP);
            tot = (row * 4) + col;
            if (tot == total)
            {
                tot -= 2;
                int countX = tot % 8;
                int countY = tot / 8 * 2;
                Grid.SetColumn(GridADP, countX);
                Grid.SetRow(GridADP, countY);
                return;
            }

            col = Grid.GetColumn(GridSelf);
            row = Grid.GetRow(GridSelf);
            tot = (row * 4) + col;
            if (tot == total)
            {
                tot -= 2;
                int countX = tot % 8;
                int countY = tot / 8 * 2;
                Grid.SetColumn(GridSelf, countX);
                Grid.SetRow(GridSelf, countY);
                return;
            }

            col = Grid.GetColumn(GridPLM);
            row = Grid.GetRow(GridPLM);
            tot = (row * 4) + col;
            if (tot == total)
            {
                tot -= 2;
                int countX = tot % 8;
                int countY = tot / 8 * 2;
                Grid.SetColumn(GridPLM, countX);
                Grid.SetRow(GridPLM, countY);
                return;
            }

            col = Grid.GetColumn(GridOracle);
            row = Grid.GetRow(GridOracle);
            tot = (row * 4) + col;
            if (tot == total)
            {
                tot -= 2;
                int countX = tot % 8;
                int countY = tot / 8 * 2;
                Grid.SetColumn(GridOracle, countX);
                Grid.SetRow(GridOracle, countY);
                return;
            }

            col = Grid.GetColumn(GridOffice);
            row = Grid.GetRow(GridOffice);
            tot = (row * 4) + col;
            if (tot == total)
            {
                tot -= 2;
                int countX = tot % 8;
                int countY = tot / 8 * 2;
                Grid.SetColumn(GridOffice, countX);
                Grid.SetRow(GridOffice, countY);
                return;
            }

            col = Grid.GetColumn(GridInstall);
            row = Grid.GetRow(GridInstall);
            tot = (row * 4) + col;
            if (tot == total)
            {
                tot -= 2;
                int countX = tot % 8;
                int countY = tot / 8 * 2;
                Grid.SetColumn(GridInstall, countX);
                Grid.SetRow(GridInstall, countY);
                return;
            }

            col = Grid.GetColumn(GridVisual);
            row = Grid.GetRow(GridVisual);
            tot = (row * 4) + col;
            if (tot == total)
            {
                tot -= 2;
                int countX = tot % 8;
                int countY = tot / 8 * 2;
                Grid.SetColumn(GridVisual, countX);
                Grid.SetRow(GridVisual, countY);
                return;
            }

            col = Grid.GetColumn(GridCalculator);
            row = Grid.GetRow(GridCalculator);
            tot = (row * 4) + col;
            if (tot == total)
            {
                tot -= 2;
                int countX = tot % 8;
                int countY = tot / 8 * 2;
                Grid.SetColumn(GridCalculator, countX);
                Grid.SetRow(GridCalculator, countY);
                return;
            }

            col = Grid.GetColumn(GridNotepad);
            row = Grid.GetRow(GridNotepad);
            tot = (row * 4) + col;
            if (tot == total)
            {
                tot -= 2;
                int countX = tot % 8;
                int countY = tot / 8 * 2;
                Grid.SetColumn(GridNotepad, countX);
                Grid.SetRow(GridNotepad, countY);
                return;
            }

            col = Grid.GetColumn(GridChrome);
            row = Grid.GetRow(GridChrome);
            tot = (row * 4) + col;
            if (tot == total)
            {
                tot -= 2;
                int countX = tot % 8;
                int countY = tot / 8 * 2;
                Grid.SetColumn(GridChrome, countX);
                Grid.SetRow(GridChrome, countY);
                return;
            }

            col = Grid.GetColumn(GridInternet);
            row = Grid.GetRow(GridInternet);
            tot = (row * 4) + col;
            if (tot == total)
            {
                tot -= 2;
                int countX = tot % 8;
                int countY = tot / 8 * 2;
                Grid.SetColumn(GridInternet, countX);
                Grid.SetRow(GridInternet, countY);
                return;
            }
        }
    }
}