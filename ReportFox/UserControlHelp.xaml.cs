using System.Windows;
using System.Windows.Controls;
using System.Diagnostics;

namespace ReportFox
{
    /// <summary>
    /// Логика взаимодействия для UserControlHelp.xaml
    /// </summary>
    public partial class UserControlHelp : UserControl
    {
        public UserControlHelp()
        {
            InitializeComponent();
        }

        private void Help(object sender, RoutedEventArgs e)
        {
            Process.Start(@"\HelpReportFox.chm");
        }
    }
}
