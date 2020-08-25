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

namespace ReportFox
{
    /// <summary>
    /// Логика взаимодействия для UserControlWhileList.xaml
    /// </summary>
    public partial class UserControlWhileList : UserControl
    {
        SoftListEntities1 SoftListDB = new SoftListEntities1(); //База данных SoftList
        List<SoftList> softLists = new List<SoftList>(); // Таблица SoftLists
        List<List_ID> list_ID = new List<List_ID>(); // Таблица List_ID

        public UserControlWhileList()
        {
            InitializeComponent();

            softLists.Clear();
            GridWhiteList.ItemsSource = null;
            GridWhiteList.Items.Clear();
            foreach (var item in SoftListDB.SoftList.ToList())
            {
                if (item.List_ID == 2)
                    softLists.Add(item);
            }
            GridWhiteList.ItemsSource = softLists;

        }
        public void LoadData(string NamePO)//поиск Имя ПО в базе
        {
            if (NamePO == "")//пустой запрос
                softLists = SoftListDB.SoftList.ToList();//вывод всех значений
            foreach (var item in SoftListDB.SoftList.ToList())
            {
                if (item.Soft_Name == NamePO)
                    softLists.Add(item);
            }
        }

        private void Find_PO(object sender, RoutedEventArgs e)
        {
            softLists.Clear();
            GridWhiteList.ItemsSource = null;
            GridWhiteList.Items.Clear();
            LoadData(Nametext.Text);//функция поиска
            GridWhiteList.ItemsSource = softLists;//загрузка результатов поиска
        }

    }
}
