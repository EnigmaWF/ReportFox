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
    /// Логика взаимодействия для UserControlSearchPO.xaml
    /// </summary>
    public partial class UserControlSearchPO : UserControl
    {
        InventoryEntities1 Inventory = new InventoryEntities1(); //База данных Inventory
        List<Hardware> hardwares = new List<Hardware>(); // Таблица Hardware 
        List<Software> softwares = new List<Software>(); // Таблица Software
        List<Users> users = new List<Users>(); //Таблица User
        public UserControlSearchPO()
        {
            InitializeComponent();
            dataInventory.ItemsSource = Inventory.Software.ToList();//выгрузка
        }

        public void LoadData(string NamePO)//поиск ID АРМ в базе
        {
            if (NamePO == "")
                softwares = Inventory.Software.ToList();
            foreach (var item in Inventory.Software.ToList())
            {
                if (item.Name == NamePO)
                    softwares.Add(item);
            }

        }

        private void FindPO(object sender, RoutedEventArgs e)
        {
            softwares.Clear();
            dataInventory.ItemsSource = null;
            dataInventory.Items.Clear();
            LoadData(Nametext.Text);
            dataInventory.ItemsSource = softwares;
        }

        private string Find(Software softwares)
        {//поиск ФИО в Users для подробного вывода
            string fio = "";
            foreach (var SoftItem in Inventory.Software.ToList())
            {
                if (SoftItem.Hardware_ID == softwares.Hardware_ID)
                {
                    foreach (var Harditem in Inventory.Hardware.ToList())
                    {
                        if (SoftItem.Hardware_ID == Harditem.ID)
                        {
                            foreach (var UsersItem in Inventory.Users.ToList())
                            {
                                if (UsersItem.UserID == Harditem.UserID)
                                {
                                    fio = UsersItem.Surname + " " + UsersItem.Name + " " + UsersItem.Patronymic;
                                }
                            }
                        }
                    }
                }
            }
            return fio;
        }

        private void MouseDouble(object sender, MouseButtonEventArgs e)
        {//Вывод подробной информации при нажатии на элемент в DataGrid
            Software softwares = dataInventory.SelectedItem as Software;
            //поиск ФИО в Users для подробного вывода
            string fio = Find(softwares);
            MessageBox.Show("\n ID АРМ: " + softwares.Hardware_ID + "\n ФИО: " + fio +
                "\n Название ПО: " + softwares.Name + "\n Версия: " + softwares.Version + "\n Расположение: " + softwares.Folder +
                "\n Дата установки: " + softwares.Installdate + "\n Дата инвентаризации: " + softwares.Lastdate, "Подробно",
                MessageBoxButton.OK, MessageBoxImage.Information); //id hardware_id(id ARM) Name (название по) (версия) расплопожение дата установки дата инвентаризации
        }
        private void TextBoxFind(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                FindPO(Nametext, e);
            }
        }

        private void Report(object sender, RoutedEventArgs e)
        {
            try
            {
                //Создание Excel отчёта

                Excel.Application excel = new Excel.Application();
                excel.Visible = true;

                Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
                Worksheet sheet1 = (Worksheet)workbook.Sheets[1];//Листы
                
                sheet1.Cells[1, 1] = "Код пользователя";
                sheet1.Cells[1, 2] = "ФИО";
                sheet1.Cells[1, 6] = "Дата проведения инвентаризации";
                
                for (int j = 0; j < dataInventory.Items.Count-1; j++)
                {
                    foreach (var SoftItem in Inventory.Software.ToList())
                    {
                        if (SoftItem.Hardware_ID == softwares[j].Hardware_ID)
                        {
                            foreach (var Harditem in Inventory.Hardware.ToList())
                            {
                                if (SoftItem.Hardware_ID == Harditem.ID)
                                {
                                    foreach (var UsersItem in Inventory.Users.ToList())
                                    {
                                        if (UsersItem.UserID == Harditem.UserID)
                                        {
                                            sheet1.Cells[j + 2, 1].Value = UsersItem.UserID;//Код пользователя
                                            sheet1.Cells[j + 2, 2].Value = UsersItem.Surname + " " + UsersItem.Name + " " + UsersItem.Patronymic;//ФИО
                                            sheet1.Cells[j + 2, 6].Value = SoftItem.Lastdate;//Дата инвентаризации 
                                            sheet1.Cells.EntireColumn.AutoFit();
                                            sheet1.Cells.EntireRow.AutoFit();
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                for (int j = 0; j < dataInventory.Columns.Count; j++) //Заголовки
                {
                    Range range = (Range)sheet1.Cells[1, j + 3];
                    range.Value2 = dataInventory.Columns[j].Header; //запись заголовка
                }

                //заполнение ячеек
                for (int i = 0; i < dataInventory.Items.Count; i++)
                {
                    for (int j = 0; j < dataInventory.Items.Count; j++)
                    {
                        TextBlock b = dataInventory.Columns[i].GetCellContent(dataInventory.Items[j]) as TextBlock;
                        Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[j + 2, i + 3];
                        range.EntireColumn.AutoFit();
                        range.EntireRow.AutoFit();
                        range.Value2 = b.Text;
                    }
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Ошибка создания отчета: \n" + ex, "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
