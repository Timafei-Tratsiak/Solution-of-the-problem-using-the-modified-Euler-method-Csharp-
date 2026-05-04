using System;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Курсовой_проект.ViewModels;

namespace Курсовой_проект.Views
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainViewModel();
            this.Closing += MainWindow_Closing;
        }

        private void MainWindow_Closing(object sender, CancelEventArgs e)
        {
            var vm = DataContext as MainViewModel;
            vm?.SaveSettingsOnExit();
        }

        private void HelpExecuted(object sender, ExecutedRoutedEventArgs e)
        {
            string helpPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "MyHelp.chm");

            if (System.IO.File.Exists(helpPath))
                System.Diagnostics.Process.Start(helpPath);
            else
                MessageBox.Show("Файл справки не найден.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
        }

        // Обработчик выбора функции из выпадающего списка (ДОБАВЛЯЕТ, а не заменяет)
        private void CmbFunctions_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbFunctions.SelectedItem != null)
            {
                var selectedItem = cmbFunctions.SelectedItem as ComboBoxItem;
                if (selectedItem != null && selectedItem.Tag != null)
                {
                    string functionToAdd = selectedItem.Tag.ToString();

                    // Получаем текущий текст и позицию курсора
                    string currentText = txtFunction.Text;
                    int cursorPosition = txtFunction.SelectionStart;

                    // Вставляем функцию в позицию курсора
                    string newText = currentText.Insert(cursorPosition, functionToAdd);
                    txtFunction.Text = newText;

                    // Восстанавливаем позицию курсора после вставки
                    txtFunction.SelectionStart = cursorPosition + functionToAdd.Length;
                    txtFunction.SelectionLength = 0;
                    txtFunction.Focus();

                    // Сбрасываем выделение
                    cmbFunctions.SelectedIndex = -1;
                }
            }
        }
    }
}