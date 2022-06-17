using System.Windows;
using System.Windows.Controls;

namespace StegoLine.Pages.Home {
    /// <summary>
    /// Interaction logic for HomePage.xaml
    /// </summary>
    public partial class HomePage: Page {
        public HomePage() {
            InitializeComponent();
        }

        private void ConcealPageBtn_Click(object sender, RoutedEventArgs e) {
            _ = this.NavigationService.Navigate(new Conceal.ConcealMsgPage());
        }

        private void RevealPageBtn_Click(object sender, RoutedEventArgs e) {
            _ = this.NavigationService.Navigate(new Reveal.RevealPage());
        }

        private void ChecksumPageBtn_Click(object sender, RoutedEventArgs e) {
            _ = this.NavigationService.Navigate(new Check.CheckPage());
        }


    }
}
