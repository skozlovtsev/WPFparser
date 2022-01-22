using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

using Newtonsoft.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using HtmlAgilityPack;
using DocumentFormat.OpenXml;

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
    public abstract class ReportValidator
    {
        public static List<Report> Validate(DataRow curr, DataRow prev)
        {
            List<Report> rep = new List<Report>();
            {
                for(int i = 1; i < curr.ItemArray.Length - 2; i++)
                {
                    if (curr.ItemArray.ToList()[i].ToString() != prev.ItemArray.ToList()[i].ToString())
                    {
                        rep.Add(new Report() { id = curr.ItemArray.ToList()[0].ToString(), current = curr.ItemArray.ToList()[i].ToString(), previous = prev.ItemArray.ToList()[i].ToString() }) ;
                    }
                }
            }
            return rep;
        }
    }
    public class Report
    {
        public string id { set; get; }
        public string current { set; get; }
        public string previous { set; get; }
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
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return null;
            }
        }
    }
}
