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

using Excel = Microsoft.Office.Interop.Excel; //библиотеки для работы с Excel
using Microsoft.Office.Interop.Excel;

namespace ReportFox
{
    /// <summary>
    /// Логика взаимодействия для UserControlBlackList.xaml
    /// </summary>
    public partial class UserControlBlackList : UserControl
    {
        SoftListEntities1 SoftListDB = new SoftListEntities1(); //База данных SoftList
        List<SoftList> softLists = new List<SoftList>(); // Таблица SoftLists
        List<List_ID> list_ID = new List<List_ID>(); // Таблица List_ID
        InventoryEntities1 Inventory = new InventoryEntities1(); //База данных Inventory
        List<Hardware> hardwares = new List<Hardware>(); // Таблица Hardware 
        List<Software> softwares = new List<Software>(); // Таблица Software
        List<Users> users = new List<Users>(); //Таблица User

        public UserControlBlackList()
        {
            InitializeComponent();
            softLists.Clear();
            GridBlackList.ItemsSource = null;
            GridBlackList.Items.Clear();
            GridBlack();
            GridBlackList.ItemsSource = softLists;
        }

        private void GridBlack()
        {
            foreach (var item in SoftListDB.SoftList.ToList())
            {
                if (item.List_ID == 1)
                    softLists.Add(item);
            }
        }
        public void LoadList(string NamePO)//поиск Имя ПО в базе
        {
            if (NamePO == "")//пустой запрос
                GridBlack();//вывод всех значений
            foreach (var item in SoftListDB.SoftList.ToList())
            {
                if (item.Soft_Name == NamePO)
                    softLists.Add(item);
            }
        }

        private void Find_Name(object sender, RoutedEventArgs e)
        {
            softLists.Clear();
            GridBlackList.ItemsSource = null;
            GridBlackList.Items.Clear();
            LoadList(Nametext.Text);//функция поиска
            GridBlackList.ItemsSource = softLists;//загрузка результатов поиска
        }

        private void TextBoxFind(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                Find_Name(Nametext, e);
            }
        }

        private void Report(object sender, RoutedEventArgs e)
        {
            try
            {
                int i = 1; 

                //Создание Excel отчёта
                Excel.Application excel = new Excel.Application();
                excel.Visible = true;

                Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
                Worksheet sheet1 = (Worksheet)workbook.Sheets[1];//Листы

                for (int k = 0; k < GridBlackList.Items.Count-1; k++) 
                {
                    //Поиск АРМ с ПО из чёрного списка
                    softwares.Clear();//очищаем список

                    //Запуск поиска АРМ с ПО
                    
                    //Заголовки
                    sheet1.Cells[1, 1] = "Код пользователя";
                    sheet1.Cells[1, 2] = "ФИО";

                    sheet1.Cells[1, 3] = "Код АРМ";
                    sheet1.Cells[1, 4] = "Имя АРМ";
                    sheet1.Cells[1, 5] = "Описание";

                    sheet1.Cells[1, 6] = "Название ПО";
                    sheet1.Cells[1, 7] = "Путь";
                    sheet1.Cells[1, 8] = "Дата установки";
                    sheet1.Cells[1, 9] = "Дата проведения инвентаризации";

                    //Выборка данных

                    //список компов
                    for (int j = 0; j < softLists.Count; j++)//softLists.Count
                    {
                        foreach (var item in Inventory.Software.ToList())
                        {
                            if (item.Name == softLists[j].Soft_Name)
                                softwares.Add(item);
                        }
                    }
                    //выборку данных
                    for (int j = 0; j < softwares.Count; j++)
                    { 
                        foreach (var SoftItem in Inventory.Software.ToList())
                        {
                            if (softLists[j].Soft_Name == SoftItem.Name)
                            {
                                i++;
                                foreach (var Harditem in Inventory.Hardware.ToList())//
                                {
                                    if (SoftItem.Hardware_ID == Harditem.ID)
                                    {
                                        foreach (var UsersItem in Inventory.Users.ToList())
                                        {
                                            if (UsersItem.UserID == Harditem.UserID)
                                            {
                                                sheet1.Cells[j + i, 1].Value = UsersItem.UserID;//Код пользователя
                                                sheet1.Cells[j + i, 2].Value = UsersItem.Surname + " " + UsersItem.Name + " " + UsersItem.Patronymic;//ФИО
                                                sheet1.Cells[j + i, 3].Value = Harditem.ID;//Код АРМ
                                                sheet1.Cells[j + i, 4].Value = Harditem.Name;//Код АРМ
                                                sheet1.Cells[j + i, 5].Value = Harditem.Description;//Описание
                                                sheet1.Cells[j + i, 6].Value = SoftItem.Name;//Название ПО
                                                sheet1.Cells[j + i, 7].Value = SoftItem.Folder;//Путь
                                                sheet1.Cells[j + i, 8].Value = SoftItem.Installdate;//Дата установки
                                                sheet1.Cells[j + i, 9].Value = SoftItem.Lastdate;//Дата инвентаризации 
                                                sheet1.Cells.EntireColumn.AutoFit();
                                                sheet1.Cells.EntireRow.AutoFit();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        i--;
                    }
                }
            }
            catch (Exception ex)
            {
               MessageBox.Show("Ошибка создания отчета: \n" + ex, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
