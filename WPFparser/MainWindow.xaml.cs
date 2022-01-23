using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.IO;
using System.Data;

namespace WPFparser
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// Полная информация о угразе и Отчеты(если изменения были) выводятся в новом окне Window1 или Window2 соответственно 
    /// </summary>
    public partial class MainWindow : Window
    {
        //Путь к текущей дирректории (для сохранения данных)
        public string dirPath = Directory.GetCurrentDirectory();
        //Путь к файлу
        public string filePath;
        //Адрес сайта ФСТЭК России
        readonly string link = @"https://bdu.fstec.ru";
        //Текущая страница
        public int page = 1;
        //Размер страници
        public int pageSize = 15;
        public DataTable data;
        public MainWindow()
        {
            InitializeComponent();
            filePath = $"{dirPath}\\data.xlsx";
            pathTB.Text = dirPath;
            currPage.Text = page.ToString();
            pageCount.Text = pageSize.ToString();
            /*В том случае если вайл не существует в messageBox добавляется сообщение (В ListBox messageBox выводятся сообщения и ошибки)
              В случае если фаил существует, информация (Идентификатор угрозы в формате «УБИ.ХХХ», Наименование угрозы) выводится в ListBox view*/
            if (!File.Exists(filePath))
            {
                messageBox.Items.Add($"Файла с локальной базой не существует. Для загрузки нажмите кнопку Reload");
            }
            try
            {
                data = Xlsx.ReadExcelas(filePath);
                if (data != null)
                {
                    //"Идентификатор УБИ" задается как Первичный Ключ
                    data.PrimaryKey = new DataColumn[] { data.Columns["Идентификатор УБИ"] };
                    for (int i = (page - 1) * pageSize; (i < data.Rows.Count) & (i < (page * pageSize)); i++)
                    {
                        view.Items.Add($"УБИ.{Ubi.Zeros(data.Rows[i].ItemArray.ToList()[0].ToString())}{data.Rows[i].ItemArray.ToList()[0]}   {data.Rows[i].ItemArray.ToList()[1]}");
                    }
                }
            }
            catch (Exception ex) { messageBox.Items.Add($"{ex}"); };
        }
        private void PathTB_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                FileAttributes attr = File.GetAttributes(pathTB.Text);
                if (attr == FileAttributes.Directory)
                {
                    dirPath = pathTB.Text;
                    filePath = $"{dirPath}\\data.xlsx";
                    reloadStatus.Foreground = Brushes.LightGreen;
                    reloadStatus.Text = "Путь является валидным*";
                }
                else
                {
                    reloadStatus.Foreground = Brushes.Red;
                    reloadStatus.Text = "Путь является НЕ валидным*";
                }
            }
            catch (Exception)
            {
                reloadStatus.Foreground = Brushes.Red;
                reloadStatus.Text = "Путь является НЕ валидным*";
            }
        }

        private void ReloadButton_Click(object sender, RoutedEventArgs e)
        {
            Loader.LoadFromPathTo($@"{link}{Parser.ParseLink(link)}", filePath);
            try
            {
                DataTable oldData = data;
                data = Xlsx.ReadExcelas(filePath);
                data.PrimaryKey = new DataColumn[] { data.Columns["Идентификатор УБИ"] };
                List<Report> changed = new List<Report>();
                /*Создание отчета
                  Отчет создается только в том случае если новые данные есть с чем сравнивать*/
                if (oldData != null)
                {
                    //Отчет о изменениях представленн как List<List<Report>> где каждая запись Report это одно измененное поле 
                    List<List<Report>> reports = new List<List<Report>>();
                    //Для каждой записи проверяется дата последнего изменения данных, если у нового элемента она больше, то вся строка поэлементно сравнивается
                    for (int i = 0; i < data.Rows.Count; i++)
                    {
                        try
                        {
                            if (Convert.ToInt32(oldData.Rows[i].ItemArray.ToList()[9].ToString()) < Convert.ToInt32(data.Rows[i].ItemArray.ToList()[9].ToString()))
                            {
                                List<Report> rep = new List<Report>();
                                {
                                    for (int j = 1; j < data.Rows[j].ItemArray.Length - 2; j++)
                                    {
                                        if (data.Rows[i].ItemArray.ToList()[j].ToString() != oldData.Rows[i].ItemArray.ToList()[j].ToString())
                                        {
                                            rep.Add(new Report() { id = $"УБИ.{Ubi.Zeros(data.Rows[i].ItemArray.ToList()[0].ToString())}{data.Rows[i].ItemArray.ToList()[0]}", cell = data.Columns[j].ColumnName, current = data.Rows[j].ItemArray.ToList()[j].ToString(), previous = oldData.Rows[i].ItemArray.ToList()[j].ToString() });
                                        }
                                    }
                                }
                                reports.Add(rep);
                            }
                        }
                        catch (Exception ex)
                        {
                            if (ex is IndexOutOfRangeException)
                            {
                                messageBox.Items.Add("Запись была добавленна"); 
                            }
                            else
                            {
                                messageBox.Items.Add(ex.Message);
                            }
                        }
                    }
                    //Вывод информации о успешном обновлении
                    if (reports.Count > 0)
                    {
                        messageBox.Items.Add("Успешно.");
                        messageBox.Items.Add($"{reports.Count} записи были обновленны.");
                        Window2 window = new Window2(reports);
                        window.Show();
                    }
                    else
                    {
                        messageBox.Items.Add("Успешно.");
                        messageBox.Items.Add("0 записей было обновленно.");
                    }
                }
                //Перезагрузка данных в ListBox view 
                view.Items.Clear();
                for (int i = (page - 1) * pageSize; (i < data.Rows.Count) & (i < (page * pageSize)); i++)
                {
                    view.Items.Add($"УБИ.{Ubi.Zeros(data.Rows[i].ItemArray.ToList()[0].ToString())}{data.Rows[i].ItemArray.ToList()[0]}   {data.Rows[i].ItemArray.ToList()[1]}");
                }
                reloadStatus.Foreground = Brushes.LightGreen;
                reloadStatus.Text = "Заружено успешно*";
                messageBox.Items.Add($"Данные были успешно загружены в директорию {filePath}");
            }
            //Вывод ошибок
            catch (Exception ex) { messageBox.Items.Add($"{ex}"); };
        }
        /// <summary>
        /// Для открытия полной инвормации о угрозе безопастности дважды кликните по записи в VievBox
        /// </summary>
        private void View_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            string item = view.SelectedItem.ToString();
            DataRow Drw = data.Rows.Find(Convert.ToInt32(item.Substring(4, 3)).ToString());
            Window1 window = new Window1(Drw);
            window.Show();
        }
        private void CurrPage_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                CurrPage_Validation();
            }
        }
        private void CurrPage_LostFocus(object sender, RoutedEventArgs e)
        {
            CurrPage_Validation();
        }
        /// <summary>
        /// CurrPage_Validation() метод для обработки перехода на другую страницу
        /// </summary>
        private void CurrPage_Validation()
        {
            if (int.TryParse(currPage.Text, out int value))
            {
                if (value < 1)
                {
                    page = 1;
                    currPage.Text = page.ToString();
                }
                else if (data.Rows.Count > (value - 1) * pageSize)
                {
                    page = value;
                }
                else
                {
                    page = Convert.ToInt32(Convert.ToDouble(data.Rows.Count) / Convert.ToDouble(pageSize));
                    currPage.Text = page.ToString();
                }
                view.Items.Clear();
                for (int i = (page - 1) * pageSize; (i < data.Rows.Count) & (i < (page * pageSize)); i++)
                {
                    view.Items.Add($"УБИ.{Ubi.Zeros(data.Rows[i].ItemArray.ToList()[0].ToString())}{data.Rows[i].ItemArray.ToList()[0]}   {data.Rows[i].ItemArray.ToList()[1]}");
                }
            }
            else
            {
                page = 1;
                currPage.Text = page.ToString();
                view.Items.Clear();
                for (int i = (page - 1) * pageSize; (i < data.Rows.Count) & (i < (page * pageSize)); i++)
                {
                    view.Items.Add($"УБИ.{Ubi.Zeros(data.Rows[i].ItemArray.ToList()[0].ToString())}{data.Rows[i].ItemArray.ToList()[0]}   {data.Rows[i].ItemArray.ToList()[1]}");
                }
            }
        }
        private void PageCount_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                PageSize_Validation();
            }
        }
        private void PageCount_LostFocus(object sender, RoutedEventArgs e)
        {
            PageSize_Validation();
        }
        /// <summary>
        /// PageSize_Validation() метод для обработки изменения размера(в записях) страници
        /// </summary>
        private void PageSize_Validation()
        {
            if (int.TryParse(pageCount.Text, out int value))
            {
                if (value < 1)
                {
                    pageSize = 1;
                    pageCount.Text = pageSize.ToString();
                }
                else if ((data.Rows.Count <= value * page) & (page > 1))
                {
                    if (value >= data.Rows.Count)
                    {
                        pageSize = data.Rows.Count;
                    }
                    else
                    {
                        pageSize = value;
                    }
                    pageCount.Text = pageSize.ToString();
                    page = 1;
                    currPage.Text = page.ToString();
                }
                else if (data.Rows.Count > value * page)
                {
                    pageSize = value;
                }

                else
                {
                    pageSize = data.Rows.Count;
                    pageCount.Text = pageSize.ToString();
                }
                view.Items.Clear();
                for (int i = (page - 1) * pageSize; (i < data.Rows.Count) & (i < (page * pageSize)); i++)
                {
                    view.Items.Add($"УБИ.{Ubi.Zeros(data.Rows[i].ItemArray.ToList()[0].ToString())}{data.Rows[i].ItemArray.ToList()[0]}   {data.Rows[i].ItemArray.ToList()[1]}");
                }
            }
            else
            {
                pageSize = 15;
                pageCount.Text = pageSize.ToString();
                view.Items.Clear();
                for (int i = (page - 1) * pageSize; (i < data.Rows.Count) & (i < (page * pageSize)); i++)
                {
                    view.Items.Add($"УБИ.{Ubi.Zeros(data.Rows[i].ItemArray.ToList()[0].ToString())}{data.Rows[i].ItemArray.ToList()[0]}   {data.Rows[i].ItemArray.ToList()[1]}");
                }
            }
        }
        private void Next_Click(object sender, RoutedEventArgs e)
        {
            currPage.Text = (Convert.ToInt32(currPage.Text) + 1).ToString();
            CurrPage_Validation();
        }
        private void Nast_Click(object sender, RoutedEventArgs e)
        {
            currPage.Text = (Convert.ToInt32(currPage.Text) - 1).ToString();
            CurrPage_Validation();
        }
    }
}
