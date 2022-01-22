using System;
using System.Data;
using System.Linq;
using System.Windows;

namespace WPFparser
{
    /// <summary>
    /// Логика взаимодействия для Window1.xaml
    /// </summary>
    public partial class Window1 : Window
    {
        public Window1(DataRow data)
        {
            InitializeComponent();
            id.Text = $"УБИ.{Ubi.Zeros(data.ItemArray.ToList()[0].ToString())}{data.ItemArray.ToList()[0]}";
            fullName.Text = data.ItemArray.ToList()[1].ToString();
            fullInformation.Text = data.ItemArray.ToList()[2].ToString();
            source.Text = data.ItemArray.ToList()[3].ToString();
            confedencial.Text = data.ItemArray.ToList()[4].ToString();
            blank1.Text = data.ItemArray.ToList()[5].ToString() == "1" ? "Да" : "Нет";
            blank2.Text = data.ItemArray.ToList()[6].ToString() == "1" ? "Да" : "Нет";
            blank3.Text = data.ItemArray.ToList()[7].ToString() == "1" ? "Да" : "Нет";
            addDate.Text = DateTime.FromOADate(Convert.ToDouble(data.ItemArray.ToList()[8].ToString())).ToShortDateString();
            updateDate.Text = DateTime.FromOADate(Convert.ToDouble(data.ItemArray.ToList()[9].ToString())).ToShortDateString();
        }
    }
}
