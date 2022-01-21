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
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
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
        public DataTable data;
        public MainWindow()
        {
            InitializeComponent();
            filePath = $"{dirPath}\\data.xlsx";
            Parser parser = new Parser(link);
            Loader loader = new Loader();
            Xlsx xlsx = new Xlsx();
            pathTB.Text = dirPath;
            if (!File.Exists(filePath))
            {
                loader.LoadFromPathTo($@"{link}{parser.ParseLink()}", filePath);
                messageBox.Items.Add($"Данные были загружены в директорию {filePath}");
            }
            try
            {
                data = xlsx.ReadExcelas(filePath);
                data.Rows[0].Delete();
                for (int i = 0; i < data.Rows.Count; i++)
                {
                    view.Items.Add($"УБИ.{Zeros(data.Rows[i].ItemArray.ToList()[0].ToString())}{data.Rows[i].ItemArray.ToList()[0]}   {data.Rows[i].ItemArray.ToList()[1]}");
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
            Loader loader = new Loader();
            Xlsx xlsx = new Xlsx();
            loader.LoadFromPathTo($@"{link}{parser.ParseLink()}", filePath);
            try
            {
                DataTable data = xlsx.ReadExcelas(filePath);
                for (int i = 1; i < data.Rows.Count; i++)
                {
                    view.Items.Add($"УБИ.{Zeros(data.Rows[i].ItemArray.ToList()[0].ToString())}{data.Rows[i].ItemArray.ToList()[0]}   {data.Rows[i].ItemArray.ToList()[1]}");
                }
                reloadStatus.Foreground = Brushes.LightGreen;
                reloadStatus.Text = "Заружено успешно*";
                messageBox.Items.Add("Заружено успешно");
            }
            catch (Exception ex) { messageBox.Items.Add($"{ex}"); };
        }
        static string Zeros(string a)
        {
            switch (a.Length)
            {
                case 2:
                    return "0";
                case 1:
                    return "00";
                default:
                    return "";
            }
        }

        private void view_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            string item = view.SelectedItem.ToString();
            data.PrimaryKey = new DataColumn[] { data.Columns["Идентификатор УБИ"] };
            DataRow Drw = data.Rows.Find(Convert.ToInt32(item.Substring(4, 3)).ToString());
            Window1 window = new Window1(Drw);
            window.Show();
        }

        private void reportBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {

        }
    }
    public class Repotr
    {

    }
    public class Parser
    {
        public string Url;
        public Parser(string url)
        {
            Url = url;
        }
        public object ParseLink()
        //Парсинг ссылки на xlsx файл
        {
            try
            {
                HtmlWeb hweb = new HtmlWeb();
                HtmlDocument hdoc = hweb.Load($@"{Url}/threat");
                HtmlNode nodes = hdoc.DocumentNode;

                string table = nodes.SelectSingleNode("//*[@id='wrap']/div[3]/div/div[1]/div[2]/div/p/a[@href]").Attributes["href"].Value;
                return table;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); return ex.Message; }
        }
    }

    public class Loader
    {
        public object LoadFromPathTo(string link, string dir)
        //Загрузка файла по ссылке
        {
            try
            {
                using (var client = new System.Net.WebClient())
                {
                    client.DownloadFile($@"{link}", dir);
                }
                return 0;
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); return ex.Message; }
        }

    }

    public class Xlsx
    {
        public DataTable ReadExcelas(string path)
        {
            try
            {
                DataTable dtTable = new DataTable();
                //Lets open the existing excel file and read through its content . Open the excel using openxml sdk
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(path, false))
                {
                    //create the object for workbook part  
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    Sheets thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();

                    //using for each loop to get the sheet from the sheetcollection  
                    foreach (Sheet thesheet in thesheetcollection.OfType<Sheet>())
                    {
                        //statement to get the worksheet object by using the sheet id  
                        Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;

                        SheetData thesheetdata = theWorksheet.GetFirstChild<SheetData>();



                        for (int rCnt = 0; rCnt < thesheetdata.ChildElements.Count(); rCnt++)
                        {
                            List<string> rowList = new List<string>();
                            for (int rCnt1 = 0; rCnt1
                                < thesheetdata.ElementAt(rCnt).ChildElements.Count(); rCnt1++)
                            {

                                Cell thecurrentcell = (Cell)thesheetdata.ElementAt(rCnt).ChildElements.ElementAt(rCnt1);
                                //statement to take the integer value  
                                if (thecurrentcell.DataType != null)
                                {
                                    if (thecurrentcell.DataType == CellValues.SharedString)
                                    {
                                        int id;
                                        if (Int32.TryParse(thecurrentcell.InnerText, out id))
                                        {
                                            SharedStringItem item = workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                                            if (item.Text != null)
                                            {
                                                //first row will provide the column name.
                                                if (rCnt == 1)
                                                {
                                                    dtTable.Columns.Add(item.Text.Text);
                                                }
                                                else if (rCnt > 1)
                                                {
                                                    rowList.Add(item.Text.Text);
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (rCnt != 0)//reserved for column values
                                    {
                                        rowList.Add(thecurrentcell.InnerText);
                                    }
                                }
                            }
                            if (rCnt != 0)//reserved for column values
                                dtTable.Rows.Add(rowList.ToArray());

                        }

                    }
                    return dtTable;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }
    }

}
