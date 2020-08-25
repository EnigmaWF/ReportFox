using BeautySolutions.View.ViewModel;
using MaterialDesignThemes.Wpf;
using System;
using ReportFox;
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

namespace ReportFox
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            var menuInfo = new List<SubItem>();//создание подкатегории элемента меню Инфо о ПО
            menuInfo.Add(new SubItem("Поиск определенного ПО", new UserControlSearchPO()));
            menuInfo.Add(new SubItem("ПО на АРМ", new POARM()));
            var item0 = new ItemMenu("Инфо о ПО", menuInfo, PackIconKind.Monitor);

            var menuSoftList = new List<SubItem>();//создание подкатегорий элемента меню Списки ПО
            menuSoftList.Add(new SubItem("Белый список", new UserControlWhileList()));
            menuSoftList.Add(new SubItem("Чёрный список", new UserControlBlackList()));
            menuSoftList.Add(new SubItem("Распределение ПО", new UserControlAlocationPO()));
            var item1 = new ItemMenu("Списки ПО", menuSoftList, PackIconKind.ClipboardTextOutline);

            var item2 = new ItemHelp("Помощь", PackIconKind.Help, new UserControlHelp());

            Menu.Children.Add(new UserControlMenuItem(item0, this));//Элемент Инфо о ПО
            Menu.Children.Add(new UserControlMenuItem(item1, this));//Элемент Списки ПО
            Menu.Children.Add(new UserControlHelpItem(item2, this));//Элемент помощь
            
        }

        internal void SwitchScreen(object sender)
        {
            var screen = ((UserControl)sender);

            if (screen != null)
            {
                StackPanelMain.Children.Clear();
                StackPanelMain.Children.Add(screen);
            }
        }
    }
}
