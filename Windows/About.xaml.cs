using MahApps.Metro.Controls;
using System.Windows;

namespace StegoLine.Windows {
    /// <summary>
    /// Interaction logic for About.xaml
    /// </summary>
    public partial class About: MetroWindow {
        public About() {
            InitializeComponent();
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e) {
            this.Close();
        }
    }
}
