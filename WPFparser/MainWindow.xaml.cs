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
using System.Net.Http;
using System.Net.Http.Headers;

using DocumentFormat.OpenXml;
using HtmlAgilityPack;
using System.IO;
using System.Data;
using Newtonsoft.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WPFparser
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string dirPath = Directory.GetCurrentDirectory();
        public string filePath;
        readonly string link = @"https://bdu.fstec.ru";
        public int page = 1;
        public int pageSize = 15;
        public DataTable data;
        public MainWindow()
        {
            InitializeComponent();
            filePath = $"{dirPath}\\data.xlsx";
            Parser parser = new Parser(link);
            pathTB.Text = dirPath;
            if (!File.Exists(filePath))
            {
                Loader.LoadFromPathTo($@"{link}{parser.ParseLink()}", filePath);
                messageBox.Items.Add($"Данные были загружены в директорию {filePath}");
            }
            try
            {
                data = Xlsx.ReadExcelas(filePath);
                data.PrimaryKey = new DataColumn[] { data.Columns["Идентификатор УБИ"] };
                currPage.Text = page.ToString();
                pageCount.Text = pageSize.ToString();
                for (int i = (page - 1) * pageSize; (i < data.Rows.Count) & (i < (page * pageSize)); i++)
                {
                    view.Items.Add($"УБИ.{Ubi.Zeros(data.Rows[i].ItemArray.ToList()[0].ToString())}{data.Rows[i].ItemArray.ToList()[0]}   {data.Rows[i].ItemArray.ToList()[1]}");
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
                    reloadStatus.Foreground = Brushes.LightGreen;
                    reloadStatus.Text = "Путь является валидным*";
                }
                else
                {
                    reloadStatus.Foreground = Brushes.Red;
                    reloadStatus.Text = "Путь является НЕ валидным*";
                }
            }
            catch (Exception ex)
            {
                reloadStatus.Foreground = Brushes.Red;
                reloadStatus.Text = "Путь является НЕ валидным*";
            }
        }

        private void ReloadButton_Click(object sender, RoutedEventArgs e)
        {
            Parser parser = new Parser(link);
            Loader.LoadFromPathTo($@"{link}{parser.ParseLink()}", filePath);
            try
            {
                DataTable oldData = data;
                data = Xlsx.ReadExcelas(filePath);
                List<Report> changed = new List<Report>();
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    try
                    {
                        if (Convert.ToInt32(oldData.Rows[i].ItemArray.ToList()[9].ToString()) < Convert.ToInt32(data.Rows[i].ItemArray.ToList()[9].ToString()))
                        {
                            //validation 
                            messageBox.Items.Add(oldData.Rows[i].ItemArray.ToList()[0].ToString());
                        }
                    }
                    catch (Exception ex)
                    {
                        if (ex is IndexOutOfRangeException)
                        {
                            messageBox.Items.Add(ex.Message);
                        }
                    }
                }
                view.Items.Clear();
                for (int i = (page - 1) * pageSize; (i < data.Rows.Count) & (i < (page * pageSize)); i++)
                {
                    view.Items.Add($"УБИ.{Ubi.Zeros(data.Rows[i].ItemArray.ToList()[0].ToString())}{data.Rows[i].ItemArray.ToList()[0]}   {data.Rows[i].ItemArray.ToList()[1]}");
                }
                reloadStatus.Foreground = Brushes.LightGreen;
                reloadStatus.Text = "Заружено успешно*";
                messageBox.Items.Add("Заружено успешно");
            }
            catch (Exception ex) { messageBox.Items.Add($"{ex}"); };
        }
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
                PageCount_Validation();
            }
        }
        private void PageCount_LostFocus(object sender, RoutedEventArgs e)
        {
            PageCount_Validation();
        }
        private void PageCount_Validation()
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
