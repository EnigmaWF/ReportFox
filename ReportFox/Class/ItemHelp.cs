using MaterialDesignThemes.Wpf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace BeautySolutions.View.ViewModel
{
    public class ItemHelp
    {
        public ItemHelp(string header, PackIconKind icon, UserControl screen = null)
        {
            Header = header;
            Screen = screen;
            Icon = icon;
        }
        public string Header { get; private set; }
        public PackIconKind Icon { get; private set; }
        public UserControl Screen { get; private set; }
    }
}
