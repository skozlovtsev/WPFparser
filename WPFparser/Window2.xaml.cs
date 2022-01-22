using System.Collections.Generic;
using System.Windows;

namespace WPFparser
{
    /// <summary>
    /// Логика взаимодействия для Window2.xaml
    /// Window2 используется как шаблон для отчета о изменениях 
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
