using System.Windows;
using Курсовой_проект.ViewModels;

namespace Курсовой_проект.Views
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainViewModel();
        }
    }
}