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
using System.Windows.Shapes;

namespace WPFparser
{
    /// <summary>
    /// Логика взаимодействия для Window2.xaml
    /// </summary>
    public partial class Window2 : Window
    {
        public Window2(List<List<Report>> reports)
        {
            InitializeComponent();
            foreach(List<Report> item in reports)
            {
                foreach (Report rep in item)
                {
                    listOfChanges.Items.Add(rep);
                }
            }
        }
    }
}
