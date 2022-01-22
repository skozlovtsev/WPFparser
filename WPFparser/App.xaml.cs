using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using HtmlAgilityPack;

namespace WPFparser
{
    /// <summary>
    /// Логика взаимодействия для App.xaml
    /// </summary>
    public partial class App : Application
    {
    }
    public abstract class Ubi
    {
        public static string Zeros(string a)
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
    }
    //Отчет о изменении
    public class Report
    {
        //УБИ
        public string id { set; get; }
        //Поле которое было изменено
        public string cell { set; get; }
        //СТАЛО
        public string current { set; get; }
        //БЫЛО
        public string previous { set; get; }
    }
    public abstract class Parser
    {
        static public object ParseLink(string url)
        //Парсинг ссылки на xlsx файл
        {
            try
            {
                HtmlWeb hweb = new HtmlWeb();
                HtmlDocument hdoc = hweb.Load($@"{url}/threat");
                HtmlNode nodes = hdoc.DocumentNode;
                //Парсинг ссылки(в задании не нашел что конкретно нужно парсить, так что.. сделал как проще и логичнее) на XLSX файл с информацией о угрозах
                //Парсинг осуществляется через XPath
                string table = nodes.SelectSingleNode("//*[@id='wrap']/div[3]/div/div[1]/div[2]/div/p/a[@href]").Attributes["href"].Value;
                return table;
            }
            catch (Exception ex) { return ex.Message; }
        }
    }
    public abstract class Loader
    {
        public static object LoadFromPathTo(string link, string dir)
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
    public abstract class Xlsx
    {
        //Очень не хочу это комментировать..
        //Открытие XLSX таблици в виде объекта DataTable
        public static DataTable ReadExcelas(string path)
        {
            try
            {
                DataTable dtTable = new DataTable();
                using (SpreadsheetDocument doc = SpreadsheetDocument.Open(path, false))
                {
                    WorkbookPart workbookPart = doc.WorkbookPart;
                    Sheets thesheetcollection = workbookPart.Workbook.GetFirstChild<Sheets>();
                    foreach (Sheet thesheet in thesheetcollection.OfType<Sheet>())
                    {
                        Worksheet theWorksheet = ((WorksheetPart)workbookPart.GetPartById(thesheet.Id)).Worksheet;
                        SheetData thesheetdata = theWorksheet.GetFirstChild<SheetData>();
                        for (int rCnt = 0; rCnt < thesheetdata.ChildElements.Count(); rCnt++)
                        {
                            List<string> rowList = new List<string>();
                            for (int rCnt1 = 0; rCnt1
                                < thesheetdata.ElementAt(rCnt).ChildElements.Count(); rCnt1++)
                            {

                                Cell thecurrentcell = (Cell)thesheetdata.ElementAt(rCnt).ChildElements.ElementAt(rCnt1);
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
                                    if (rCnt != 0)
                                    {
                                        rowList.Add(thecurrentcell.InnerText);
                                    }
                                }
                            }
                            if (rCnt != 0)
                                dtTable.Rows.Add(rowList.ToArray());
                        }
                    }
                    dtTable.Rows[0].Delete();
                    return dtTable;
                }
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}
