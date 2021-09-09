using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.Diagnostics;

namespace WpfExcelApp1
{
    class ExcelMethodClass
    {
        public static void actionTwo()
        {
            try
            {
                Excel.Application excel = new Excel.Application();
                // open the concrete file
                Excel.Workbook excelWorkbook = excel.Workbooks.Open(MainWindow.nameFile);
                // select worksheet
                Excel._Worksheet excelWorkbookWorksheet = excelWorkbook.Sheets[1];

                //Последняя строка
                int LastRow = excelWorkbook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                //Последняя колонка
                int LastColumn = excelWorkbook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                HashSet<string> labelsList = new HashSet<string>(); // список Labels
                Dictionary<string, Dictionary<string, uint>> dict = new Dictionary<string, Dictionary<string, uint>>(); //словарь результатов

                for (int row = 2; CheckEnd(excelWorkbookWorksheet, row); row++) // перебор строк
                {
                    if (excelWorkbookWorksheet.Cells[row, 6].Value2 != null && excelWorkbookWorksheet.Cells[row, 7].Value2 != null) // проверка наличия данных в ячейках
                    {
                        string memberCell = excelWorkbookWorksheet.Cells[row, 6].Value2.ToString(); // получение текста ячеек
                        string labelCell = excelWorkbookWorksheet.Cells[row, 7].Value2.ToString();
                        if (memberCell != string.Empty && labelCell != string.Empty) // проверка наличия данных в ячейках
                        {
                            string[] members = memberCell.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries); // разделение текста
                            string[] labels = labelCell.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                            foreach (string member in members) // перебор Members
                            {
                                if (!dict.ContainsKey(member)) // не содержит member
                                {
                                    dict.Add(member, new Dictionary<string, uint>()); // добавить member в словарь
                                }
                                foreach (string label in labels) // перебор Labels
                                {
                                    if (!dict[member].ContainsKey(label)) // member не содержит label
                                    {
                                        dict[member].Add(label, 1); // // добавить label в словарь member
                                    }
                                    else
                                    {
                                        dict[member][label]++; //увеличить количество
                                    }
                                    if (!labelsList.Contains(label)) // не содержится в списке
                                    {
                                        labelsList.Add(label); // добавить в список
                                    }
                                }
                            }
                        }
                    }
                }

                excelWorkbook.Close(false); // закрыть книгу

                Excel.Application excelTwo = new Excel.Application();
                excelTwo.SheetsInNewWorkbook = 1; // количество листов в новой книге
                excelWorkbook = excelTwo.Workbooks.Add(); // создание книги
                excelWorkbookWorksheet = (Excel.Worksheet)excelWorkbook.Worksheets.get_Item(1); // получение листа
                excelWorkbookWorksheet.Name = "Result"; // название листа

                string[] labels2 = labelsList.ToArray(); // конвертация в массив
                uint[] sum = new uint[labels2.Length]; // суммы по каждому label 
                int row2 = 2; // номер строки

                foreach (KeyValuePair<string, Dictionary<string, uint>> item in dict) // перебор member
                {
                    excelWorkbookWorksheet.Cells[row2, 1] = item.Key; // текст member
                    for (int column = 0; column < labels2.Length; column++) // перебор label
                    {
                        if (item.Value.ContainsKey(labels2[column])) // label содержится в member
                        {
                            excelWorkbookWorksheet.Cells[row2, column + 2] = item.Value[labels2[column]]; // количество
                            sum[column] += item.Value[labels2[column]]; // суммирование
                        }
                        else
                        {
                            excelWorkbookWorksheet.Cells[row2, column + 2] = 0;
                        }
                    }
                    row2++;
                }
                for (int column = 0; column < labels2.Length; column++) // перебор label
                {
                    excelWorkbookWorksheet.Cells[1, column + 2] = labels2[column]; // текст label
                    excelWorkbookWorksheet.Cells[dict.Count + 2, column + 2] = sum[column]; // сумма по label
                }

                // Диалог сохранение файла
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                if (saveFileDialog.ShowDialog() == true)
                {
                    excelWorkbook.SaveAs(saveFileDialog.FileName);
                    excelWorkbook.Close(false);
                }

                excel.Quit();

                if (excel != null)
                {
                    Process[] pProcess;
                    pProcess = Process.GetProcessesByName("Excel");
                    pProcess[0].Kill();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка");
            }
        }

        public static void actionOne()
        {
            try
            {
                Excel.Application excel = new Excel.Application();
                // open the concrete file
                Excel.Workbook excelWorkbook = excel.Workbooks.Open(MainWindow.nameFile);
                // select worksheet
                Excel._Worksheet excelWorkbookWorksheet = excelWorkbook.Sheets[1];

                //Последняя строка
                int LastRow = excelWorkbook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;
                //Последняя колонка
                int LastColumn = excelWorkbook.Sheets[1].Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Column;

                //Удаление строк
                for (int i = LastRow; i >= 2; i--)
                {
                    if (excelWorkbookWorksheet.Cells[i, 1].Text.ToString() != @"Done 🎉")
                    {
                        excelWorkbookWorksheet.Rows[i].Delete(XlDeleteShiftDirection.xlShiftUp);
                    }
                }

                //Удаление столбцов
                for (int i = LastColumn; i >= 1; i--)
                {
                    if (excelWorkbookWorksheet.Cells[1, i].Text.ToString() == "Card URL")
                    {
                        excelWorkbookWorksheet.Columns[i].Delete();
                    }
                    if (excelWorkbookWorksheet.Cells[1, i].Text.ToString() == "Card #")
                    {
                        excelWorkbookWorksheet.Columns[i].Delete();
                    }
                    if (excelWorkbookWorksheet.Cells[1, i].Text.ToString() == "Points")
                    {
                        excelWorkbookWorksheet.Columns[i].Delete();
                    }
                }

                //Выравнивание для всех ячеек
                for (int i = LastColumn; i >= 1; i--)
                {
                    for (int j = LastRow; j >= 1; j--)
                    {
                        excelWorkbookWorksheet.Cells[i, j].VerticalAlignment = XlHAlign.xlHAlignCenter;
                        excelWorkbookWorksheet.Cells[i, j].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        excelWorkbookWorksheet.Cells[i, j].Style.WrapText = true;
                    }
                }

                //Выравнивание для одного столбца
                for (int i = LastColumn; i >= 1; i--)
                {
                    if (excelWorkbookWorksheet.Cells[1, i].Text.ToString() == "Description")
                    {
                        for (int j = LastRow; j >= 1; j--)
                        {
                            excelWorkbookWorksheet.Rows[j].HorizontalAlignment = XlHAlign.xlHAlignCenter;
                            excelWorkbookWorksheet.Cells[i, j].VerticalAlignment = XlHAlign.xlHAlignDistributed;
                            excelWorkbookWorksheet.Cells[i, j].Style.WrapText = true;
                        }
                    }
                }

                //Полужирные заголовки
                for (int i = LastColumn; i >= 1; i--)
                {
                    excelWorkbookWorksheet.Cells[1, i].Font.Bold = true;
                }

                //Выравнивание по объёму текста
                for (int i = LastRow; i >= 1; i--)
                {
                    for (int j = LastColumn; j >= 1; j--)
                    {

                        excelWorkbookWorksheet.Cells[j, i].Style.WrapText = true;
                    }
                }

                // Диалог сохранение файла
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
                saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                if (saveFileDialog.ShowDialog() == true)
                {
                    excelWorkbook.SaveAs(saveFileDialog.FileName);
                    excelWorkbook.Close(false);
                }

                excel.Quit();

                if (excel != null)
                {
                    Process[] pProcess;
                    pProcess = Process.GetProcessesByName("Excel");
                    pProcess[0].Kill();
                }

            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка");
            }
        }

        public static bool CheckEnd(_Worksheet excelWorkbookWorksheet, int row)
        {
            for (int column = 1; column <= 9; column++)
            {
                object cellValue = excelWorkbookWorksheet.Cells[row, column].Value2;
                if (cellValue != null && cellValue.ToString() != string.Empty)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
