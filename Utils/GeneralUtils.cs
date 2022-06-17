using System;
using System.Security.Cryptography;
using System.Text;
using System.Windows;

namespace StegoLine.Utils {
    public class GeneralUtils {

        public static Encoding GetEncoding(int codepage) {
            try {
                return Encoding.GetEncoding(codepage);
            }
            catch (Exception) {
                //Window Win = Application.Current.MainWindow;
                //MessageBox.Show("Your system does not support Windows-1251 code page!", "Error", MessageBoxButton.OK, MessageBoxImage.Stop);
                ShowErrorBox(Application.Current.Resources["CodepageError"].ToString());
                Application.Current.Shutdown();
                return Encoding.Default;
            }
        }
        public static string HashCode(string msg, Encoding MsgEncoding) {
            SHA384 SHA = SHA384.Create();
            byte[] Hash = SHA.ComputeHash(MsgEncoding.GetBytes(msg));
            return Convert.ToHexString(Hash);
        }

        public static void ShowErrorBox(string? ExMsg) {
            _ = MessageBox.Show(
                    ExMsg,
                    Application.Current.Resources["ErrorBoxHeader"].ToString(),
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
        }

        public static void ShowSuccessBox(string? Msg) {
            _ = MessageBox.Show(
                    Msg,
                    Application.Current.Resources["SuccessBoxHeader"].ToString(),
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
        }

        public static void ShowInfoBox(string? Msg) {
            _ = MessageBox.Show(
                    Msg,
                    Application.Current.Resources["InfoBoxHeader"].ToString(),
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
        }
    }
}
