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
    /// Логика взаимодействия для UserControlAlocationPO.xaml
    /// </summary>
    public partial class UserControlAlocationPO : UserControl
    {
        SoftListEntities1 SoftListDB = new SoftListEntities1(); //База данных SoftList
        List<SoftList> softListsW = new List<SoftList>(); // Таблица SoftLists белый список
        List<SoftList> softListsB = new List<SoftList>(); // Таблица SoftLists чёный список
        List<SoftList> softLists = new List<SoftList>(); // Таблица SoftLists распределение

        InventoryEntities1 Inventory = new InventoryEntities1(); //База данных Inventory
        List<Software> softwares = new List<Software>(); // Таблица Software

        public UserControlAlocationPO()
        {
            InitializeComponent();
            //выборка рарешенного ПО в GridWhiteList
            UploadWhite();
            //выборка несанкционированного ПО в GridBlackList
            UploadBlack();
            //выборка по для распределения
            UploadAllocation();
        }

        public void UploadBlack()
        {//выборка рарешенного ПО в GridWhiteList
            softListsB.Clear();
            GridBlackList.ItemsSource = null;
            GridBlackList.Items.Clear();
            foreach (var item in SoftListDB.SoftList.ToList())
            {//выборка по коду ListID 
                if (item.List_ID == 1)//1 соотвествует белому списку
                    softListsB.Add(item);
            }
            GridBlackList.ItemsSource = softListsB;
        }

        public void UploadWhite()
        {//выборка несанкционированного ПО в GridBlackList
            softListsW.Clear();
            GridWhiteList.ItemsSource = null;
            GridWhiteList.Items.Clear();
            foreach (var item in SoftListDB.SoftList.ToList())
            {//выборка по коду ListID 
                if (item.List_ID == 2)//2 соотвествует чёрному списку
                    softListsW.Add(item);
            }
            GridWhiteList.ItemsSource = softListsW;
        }
        public void UploadAllocation()
        {//выборка несанкционированного ПО в GridBlackList
            softLists.Clear();
            GridAllocation.ItemsSource = null;
            GridAllocation.Items.Clear();
            foreach (var item in SoftListDB.SoftList.ToList())
            {//выборка по коду ListID 
                if (item.List_ID == 0)//0 - новое ПО
                    softLists.Add(item);
            }
            GridAllocation.ItemsSource = softLists;
        }

        private void ToBlack(object sender, RoutedEventArgs e)
        {//загрузить выбранный элемент
            SoftList softList = GridWhiteList.SelectedItem as SoftList;//выбор элемента  
            SoftList AllocList = GridAllocation.SelectedItem as SoftList;//выбор элемента
            if (softList != null)
            {
                //внести изменения
                softList.List_ID = 1;
                //сохранить изменения
                SoftListDB.SaveChanges();
            }
            else if (AllocList != null)
            {
                //внести изменения
                AllocList.List_ID = 1;
                //сохранить изменения
                SoftListDB.SaveChanges();
            }
            else
            {
                MessageBox.Show("Выберите элемент для переноса в чёрный список", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            //сохранить изменения
            SoftListDB.SaveChanges();
            UploadBlack();
            UploadWhite();
            UploadAllocation();
        }

        private void ToWhite(object sender, RoutedEventArgs e)
        {//загрузить выбранный элемент
            SoftList softList = GridBlackList.SelectedItem as SoftList;//выбор элемента   
            SoftList AllocList = GridAllocation.SelectedItem as SoftList;//выбор элемента
            if (softList!=null)
            {
                //внести изменения
                softList.List_ID = 2;
                //сохранить изменения
                SoftListDB.SaveChanges();
            }
            else if (AllocList!=null)
            {
                //внести изменения
                AllocList.List_ID = 2;
                //сохранить изменения
                SoftListDB.SaveChanges();
            }
            else
            {
                MessageBox.Show("Выберите элемент для переноса в белый список", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            UploadWhite();
            UploadBlack();
            UploadAllocation();
        }

        public void Update(object sender, RoutedEventArgs e)
        {
            softLists.Clear();
            GridAllocation.ItemsSource = null;
            GridAllocation.Items.Clear();


            int k = 0;//счётчик совпадений

            foreach (var item2 in Inventory.Software.ToList())
             { 
                foreach (var item in SoftListDB.SoftList.ToList())
                    {//выборка по коду ListID 
                        if (item.Soft_Name == item2.Name)
                        {
                            k = 1;//совпадение есть
                        }
                }

                if (k == 0)
                {
                    SoftList softlistss = new SoftList();
                    {
                        softlistss.Soft_Name = item2.Name;
                        softlistss.List_ID = 0;
                    }
                    SoftListDB.SoftList.Add(softlistss);
                    SoftListDB.SaveChanges();
                }
                k = 0;
             }
            UploadAllocation();
        }
    }
}
