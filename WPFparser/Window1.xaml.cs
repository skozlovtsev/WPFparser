using System;
using System.Collections.Generic;
using System.Data;
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
using System.Windows.Shapes;

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
            addDate.Text = DateTime.FromOADate(Convert.ToDouble(data.ItemArray.ToList()[8].ToString())).ToShortDateString();
            updateDate.Text = DateTime.FromOADate(Convert.ToDouble(data.ItemArray.ToList()[9].ToString())).ToShortDateString();
        }
    }
}
