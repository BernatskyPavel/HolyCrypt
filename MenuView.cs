using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace HolyCryptv3
{
    partial class MainWindow : Window
    {
        private void MenuBtn_Click(object sender, RoutedEventArgs e) {
            Button button = (Button)sender;
            switch (button.Name) {
                case "ConcealMsgBtn":
                    TabMenu.SelectedIndex = 1;
                    break;
                case "RevealMsgBtn":
                    TabMenu.SelectedIndex = 2;
                    break;
                default:
                    break;
            }
        }
    }
}
