using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using Excel = Microsoft.Office.Interop.Excel;
using MessageBox = System.Windows.MessageBox;
using Microsoft.Office.Interop.Excel;
using Window = System.Windows.Window;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;

namespace WpfExcelApp1
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // Имя открытого файла
        public static string nameFile = "";
        public MainWindow()
        {
            InitializeComponent();
            textBox1.Text = "Выберите файл для работы";
        }

        // Кнопка "открыть файл"
        private void OpenFileButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                openFileDialog.ShowDialog();

                string filename = openFileDialog.FileName;
                nameFile = filename;

                textBox1.Text = "Выберите вариант работы с файлом";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка");
            }
        }

        // Вариант №1
        private void oneAction(object sender, RoutedEventArgs e)
        {
            textBox1.Text = "Ожидайте завершения работы программы...";

            ExcelMethodClass.actionOne();
        }

        // Вариант №2
        private void twoAction(object sender, RoutedEventArgs e)
        {
            textBox1.Text = "Ожидайте завершения работы программы...";

            ExcelMethodClass.actionTwo();
        }
    }
}
