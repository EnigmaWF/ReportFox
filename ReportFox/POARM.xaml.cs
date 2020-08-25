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
using System.Data.SqlClient;
using System.Data;
using System.IO;
using Microsoft.Win32;

using Excel = Microsoft.Office.Interop.Excel; //библиотеки для работы с Excel
using Microsoft.Office.Interop.Excel;

namespace ReportFox
{
    /// <summary>
    /// Логика взаимодействия для POARM.xaml
    /// </summary>
    public partial class POARM : UserControl
    {
        
        InventoryEntities1 Inventory = new InventoryEntities1(); //База данных Inventory
        List<Hardware> hardwares = new List<Hardware>(); // Таблица Hardware 
        List<Software> softwares = new List<Software>(); // Таблица Software
        List<Users> users = new List<Users>(); //Таблица User

        public POARM()
        {
            InitializeComponent();
            dataInventory.ItemsSource = Inventory.Software.ToList();//загрузка данных в DataGrid dataInventory
        }

        public void LoadData(string idarm)//поиск ID АРМ в базе
        {
            if (idarm == "")//пустой запрос
                softwares = Inventory.Software.ToList();//вывод всех значений
            foreach (var item in Inventory.Software.ToList())
            {
                if (item.Hardware_ID == idarm)//поиск записей с тем же ID АРМ
                    softwares.Add(item);//добавление в список
            }
        }

        private void FindID(object sender, RoutedEventArgs e)//поиск
        {
            softwares.Clear();
            dataInventory.ItemsSource = null;
            dataInventory.Items.Clear();
            LoadData(IDtext.Text.ToUpper());//функция поиска
            dataInventory.ItemsSource = softwares;//загрузка результатов поиска
        }


        private void Report(object sender, RoutedEventArgs e)//object sender, RoutedEventArgs e
        {//Создание Excel отчёта
            try
            {
                Excel.Application excel = new Excel.Application();
                excel.Visible = true;

                Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
                Worksheet sheet1 = (Worksheet)workbook.Sheets[1];//Листы

                sheet1.Cells[1, 1] = "Код пользователя";
                sheet1.Cells[1, 2] = "ФИО";
                sheet1.Cells[1, 3] = "Дата проведения инвентаризации";
                sheet1.Cells[4, 1] = "Код АРМ";
                sheet1.Cells[4, 2] = "Имя АРМ";
                sheet1.Cells[4, 3] = "Описание АРМ";
                
                foreach (var SoftItem in Inventory.Software.ToList()) //поиск ФИО, код пользователя, дата инвентаризации
                {
                    if (SoftItem.Hardware_ID == softwares[2].Hardware_ID)
                    {
                        foreach (var Harditem in Inventory.Hardware.ToList())
                        {
                            if (SoftItem.Hardware_ID == Harditem.ID)
                            {
                                foreach (var UsersItem in Inventory.Users.ToList())
                                {
                                    if (UsersItem.UserID == Harditem.UserID)
                                    {
                                        sheet1.Cells[2, 2].Value = UsersItem.Surname + " " + UsersItem.Name + " " + UsersItem.Patronymic;//ФИО
                                        sheet1.Cells[2, 1].Value = UsersItem.UserID;//Код пользователя
                                        sheet1.Cells[2, 3].Value = SoftItem.Lastdate;//Дата инвентаризации 
                                        sheet1.Cells[5, 1].Value = SoftItem.Hardware_ID;//Код АРМ
                                        sheet1.Cells[5, 2].Value = Harditem.Name;//Имя АРМ
                                        sheet1.Cells[5, 3].Value = Harditem.Description;//Описание
                                        sheet1.Cells.EntireColumn.AutoFit();
                                        sheet1.Cells.EntireRow.AutoFit();
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }

                for (int j = 1; j < dataInventory.Columns.Count; j++) //Заголовки
                {
                    Range range = (Range)sheet1.Cells[7, j];
                    range.Value2 = dataInventory.Columns[j].Header; //запись заголовка
                }

                //заполнение ячеек
                for (int i = 1; i < dataInventory.Items.Count; i++)
                {
                    for (int j = 0; j < dataInventory.Items.Count; j++)
                    {
                        TextBlock b = dataInventory.Columns[i].GetCellContent(dataInventory.Items[j]) as TextBlock;
                        Microsoft.Office.Interop.Excel.Range range = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[j + 8, i];
                        range.EntireColumn.AutoFit();
                        range.EntireRow.AutoFit();
                        range.Value2 = b.Text;
                    }
                }
            }
            catch(Exception ex)
            {

            }
        }

        private string Find(Software softwares)
        {
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

        private void Grid_MouseDouble(object sender, MouseButtonEventArgs e)
        {//Вывод подробной информации при нажатии на элемент в DataGrid
            Software softwares = dataInventory.SelectedItem as Software;

            //поиск ФИО в Users для подробного вывода
            string fio = Find(softwares);
            //вывод сообщения подробно
            MessageBox.Show("\n ID АРМ: " +  softwares.Hardware_ID + "\n ФИО: " + fio +
                "\n Название ПО: " + softwares.Name + "\n Версия: " + softwares.Version + "\n Расположение: " + softwares.Folder +
                "\n Дата установки: " + softwares.Installdate + "\n Дата инвентаризации: " + softwares.Lastdate, "Подробно",
                MessageBoxButton.OK, MessageBoxImage.Information); 
        }

        private void TextBoxFind(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                FindID(IDtext, e);
            }
        }
    }
}
