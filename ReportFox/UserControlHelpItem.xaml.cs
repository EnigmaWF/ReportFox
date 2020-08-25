using BeautySolutions.View.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Diagnostics;

namespace ReportFox
{
    /// <summary>
    /// Логика взаимодействия для UserControlHelpItem.xaml
    /// </summary>
    public partial class UserControlHelpItem : UserControl
    {
        MainWindow _context;
        public UserControlHelpItem(ItemHelp itemHelp, MainWindow context)
        {
            InitializeComponent();
            _context = context;
            this.DataContext = itemHelp;
            
        }
        private void Open(object sender, RoutedEventArgs e)
        {
            Process.Start(@"C:\Users\peceh\Downloads\ReportFox\ReportFox\HelpReportFox.chm");
        }
    }
}
