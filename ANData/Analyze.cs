using System;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;
using System.Threading;
using System.Drawing;

 /*
 Данная программа предназначена для обработки исходных данных Заказчика ООО "СевКомНефтегаз"
 Алгоритм позволяет определить количество и мощность максимального числа одновременно работающих скважин
 В качестве исходных данных может использоваться сразу множество файлов с несколькими страницами
 Полезная ссылка:
 https://zennolab.com/discussion/threads/sozdanie-excel-fajlov.15797/
 */
namespace ANData
{
    internal class Analyze
    {
        async public static Task AN(CancellationToken cancellationToken)
        {
            int poryad = 0;
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("№ КП", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("№", Type.GetType("System.Int32")));
            dt.Columns.Add(new DataColumn("№ скв.", Type.GetType("System.String")));
            dt.Columns.Add(new DataColumn("Дата ввода скважины", Type.GetType("System.DateTime")));
            dt.Columns["Дата ввода скважины"].DataType = Type.GetType("System.DateTime");
            dt.Columns.Add(new DataColumn("Дата перевода в ППД", Type.GetType("System.DateTime")));
            dt.Columns.Add(new DataColumn("Мощность", Type.GetType("System.Int32"))); // Создание формы таблицы с информацией
            DateTime datein;
            DateTime dateout;
            foreach (string element in variables.directory) // Перебор элементов EXcel в выбранной директории
            {
                cancellationToken.ThrowIfCancellationRequested(); // Останавливает выполнение программы
                FileInfo fileInfo = new FileInfo(element);
                ExcelPackage filelist = new ExcelPackage(fileInfo);
                ExcelWorksheets sheets = filelist.Workbook.Worksheets; // Получение списка листов в конкретном файле
                for (int i = 1; i <= sheets.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    ExcelWorksheet sheet = sheets[i];
                    var number_kp = "КП " + sheet.Cells["A5"].Value.ToString().Split('№')[1].Split('\n')[0]; // Получение и запись номера КП в удобной форме
                    for (int k = 1; k <= 24; k++) // Формирование списка дат в удобной форме
                    {
                        try
                        {
                            datein = DateTime.FromOADate(long.Parse(sheet.Cells[11 + k, 10].Value.ToString().Trim().Split(',')[0]));
                        }
                        catch (FormatException)
                        {
                            datein = Convert.ToDateTime(sheet.Cells[11 + k, 10].Value.ToString().Trim().Split(' ')[0]);
                        }
                        
                        if (sheet.Cells[11 + k, 11].Value != null && sheet.Cells[11 + k, 11].Value.ToString().Trim() != "")
                        {
                            try
                            {
                                dateout = DateTime.FromOADate(long.Parse(sheet.Cells[11 + k, 11].Value.ToString().Trim().Split(',')[0]));
                            }
                            catch (FormatException)
                            {
                                dateout = Convert.ToDateTime(sheet.Cells[11 + k, 11].Value.ToString().Trim().Split(' ')[0]);
                            }
                            variables.dateTimeOut.Add(dateout);
                        }
                        variables.dateTimeIn.Add(datein);
                    }
                    for (int j = 1; j <= 24; j++) // Запись в таблицу даных информации из загруженных Excel файлов
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        try
                        {
                            dt.Rows.Add(new Object[] { number_kp, j, sheet.Cells[11 + j, 2].Value, DateTime.FromOADate(long.Parse(sheet.Cells[11 + j, 10].Value.ToString().Trim().Split(',')[0])),
                        sheet.Cells[11 + j, 11].Value == null || sheet.Cells[11 + j, 11].Value.ToString().Trim() == "" ? default : DateTime.FromOADate(long.Parse(sheet.Cells[11 + j, 11].Value.ToString().Trim().Split(',')[0])), Convert.ToInt32(sheet.Cells[11 + j, 12].Value) }); // Заполнение таблицы информацией с каждой КП
                        }
                        catch (FormatException)
                        {
                            dt.Rows.Add(new Object[] { number_kp, j, sheet.Cells[11 + j, 2].Value, Convert.ToDateTime(sheet.Cells[11 + j, 10].Value.ToString().Trim().Split(' ')[0]),
                        sheet.Cells[11 + j, 11].Value == null || sheet.Cells[11 + j, 11].Value.ToString().Trim() == "" ? default : Convert.ToDateTime(sheet.Cells[11 + j, 11].Value.ToString().Trim().Split(' ')[0]), Convert.ToInt32(sheet.Cells[11 + j, 12].Value) }); // Заполнение таблицы информацией с каждой КП
                        }
                    }
                    dt.Rows.Add(); // Добавление строчек для суммарных значений количества и мощности скважин по каждой КП
                    dt.Rows.Add();
                    poryad++;
                }
            }

            dt.Rows.Add(); // Добавление строчек для суммарных значений количества и мощности скважин для всех файлов
            dt.Rows.Add();
            dt.Rows.Add();

            variables.dateTimepd = variables.dateTimepd.Concat(variables.dateTimeIn).ToList();
            variables.dateTimepd = variables.dateTimepd.Concat(variables.dateTimeOut).ToList(); // Объединение спсиков дат в общий
            variables.dateTimepd.RemoveAll(item => item == null); // Удаление пустых элементов
            variables.dateTimepd.Sort(); // Сортировка списка по возрастанию
            var dateTimepd_res = variables.dateTimepd.Distinct().ToList(); // Удаление из списка повторяющихся элементов

            foreach (var item in dateTimepd_res) // Добавление столбцов для формирования таблицы с мощностями
            {
                cancellationToken.ThrowIfCancellationRequested();
                dt.Columns.Add(new DataColumn(item.ToShortDateString(), Type.GetType("System.String")));
            }

            for (int i = 1; i <= dt.Rows.Count / 26; i++) // Установка нулевого значения для суммарных строк
            {
                cancellationToken.ThrowIfCancellationRequested();
                for (int j = 0; j < dateTimepd_res.Count; j++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    dt.Rows[i * 26 - 2][j + 6] = "0";
                    dt.Rows[i * 26 - 1][j + 6] = "0";
                }
            }

            for (int j = 0; j < dateTimepd_res.Count; j++) // Установка нулевого значения для суммарных строк
            {
                cancellationToken.ThrowIfCancellationRequested();
                dt.Rows[dt.Rows.Count - 2][j + 6] = "0";
                dt.Rows[dt.Rows.Count - 1][j + 6] = "0";
            }

            for (int i = 0; i < dt.Rows.Count; i++) // Формирование таблицы мощностей
            {
                cancellationToken.ThrowIfCancellationRequested();
                var need1 = (i / 26) * 26 + 26 - 2;
                var need2 = (i / 26) * 26 + 26 - 1;
                for (int j = 0; j < dateTimepd_res.Count; j++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (i % 26 == 24 || i % 26 == 25 || i == dt.Rows.Count - 1 || i == dt.Rows.Count - 2 || i == dt.Rows.Count - 3) // Исключение для суммарных строк
                    {
                        continue;
                    }
                    if (Convert.ToDateTime(dt.Rows[i]["Дата ввода скважины"]) <= dateTimepd_res[j] && (dateTimepd_res[j] <= Convert.ToDateTime(dt.Rows[i]["Дата перевода в ППД"]) || Convert.ToDateTime(dt.Rows[i]["Дата перевода в ППД"]) <= Convert.ToDateTime("01.01.2000")))
                    {
                        dt.Rows[i][j + 6] = dt.Rows[i]["Мощность"]; // Метод записи мощности в соотвествующую ячейку при выполнении необходимых условий
                        var need3 = Convert.ToInt32(dt.Rows[need1][j + 6]); // Формулы для подсчета суммарных значений количества и мощности скважин для всех файлов
                        dt.Rows[need1][j + 6] = need3 + 1;
                        var need4 = Convert.ToInt32(dt.Rows[need2][j + 6]);
                        dt.Rows[need2][j + 6] = need4 + Convert.ToInt32(dt.Rows[i]["Мощность"]);
                        var need5 = Convert.ToInt32(dt.Rows[dt.Rows.Count - 2][j + 6]);
                        dt.Rows[dt.Rows.Count - 2][j + 6] = need5 + 1;
                        var need6 = Convert.ToInt32(dt.Rows[dt.Rows.Count - 1][j + 6]);
                        dt.Rows[dt.Rows.Count - 1][j + 6] = need6 + Convert.ToInt32(dt.Rows[i]["Мощность"]);
                    }
                }
            }

            FileInfo file = new FileInfo(variables.path + "Итог.xlsx"); // Указываем путь к файлу, куда будем записывать информацию
            if (file.Exists) 
            {
                file.Delete(); // Удаляем старый файл
                file = new FileInfo(variables.path + "Итог.xlsx"); // Заново создаем файл
            }

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Итог");  // Обращаемся к первому листу
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    for (int j = 0; j < dateTimepd_res.Count + 6; j++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        worksheet.Cells[i + 2, j + 1].Value = dt.Rows[i][j].ToString(); // Производим запись таблицы в файл
                    }
                    if (i % 26 == 24 || i % 26 == 25 || i == dt.Rows.Count - 1 || i == dt.Rows.Count - 2 || i == dt.Rows.Count - 3) // Исключение для суммарных строк
                    {
                        continue;
                    }
                    if (Convert.ToDateTime(dt.Rows[i]["Дата перевода в ППД"]) == default)
                    {
                        Console.WriteLine(dt.Rows[i]["Дата перевода в ППД"]);
                        worksheet.Cells[i + 2, 5].Value = "";
                    }
                }
                for (int j = 0; j < dateTimepd_res.Count; j++) // Запись дат в верхней строке
                {
                    worksheet.Cells[1, j + 7].Value = dateTimepd_res[j].ToString("d");
                }
                for (int j = 0; j < 5; j++) // Запись заголовков столбцов
                {
                    worksheet.Cells[1, j + 1].Value = dt.Columns[j].ColumnName;
                }
                for (int i = 1; i <= dt.Rows.Count / 26; i++) // Определение максимальных значений в строках с суммарным количеством и мощностью скважин для каждой КП в отдельности
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var x1 = 0;
                    var x2 = 0;
                    for (int j = 0; j < dateTimepd_res.Count; j++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (Convert.ToInt32(dt.Rows[i * 26 - 2][j + 6]) > x1)
                        {
                            x1 = Convert.ToInt32(dt.Rows[i * 26 - 2][j + 6]);
                        }
                        if (Convert.ToInt32(dt.Rows[i * 26 - 1][j + 6]) > x2)
                        {
                            x2 = Convert.ToInt32(dt.Rows[i * 26 - 1][j + 6]);
                        }
                    }
                    worksheet.Cells[i * 26, dateTimepd_res.Count + 7].Value = x1;
                    worksheet.Cells[i * 26 + 1, dateTimepd_res.Count + 7].Value = x2;
                    for (int j = 0; j < dateTimepd_res.Count; j++) // Выделение цветом максимальных значений
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (Convert.ToInt32(dt.Rows[i * 26 - 2][j + 6]) == x1)
                        {
                            worksheet.Cells[i * 26, j + 7].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[i * 26, j + 7].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                        }
                        if (Convert.ToInt32(dt.Rows[i * 26 - 1][j + 6]) == x2)
                        {
                            worksheet.Cells[i * 26 + 1, j + 7].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[i * 26 + 1, j + 7].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                        }
                    }
                }
                var x3 = 0;
                var x4 = 0;
                for (int j = 0; j < dateTimepd_res.Count; j++) // Определение максимальных значений в строках с суммарным количеством и мощностью скважин для всех КП
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (Convert.ToInt32(dt.Rows[dt.Rows.Count - 2][j + 6]) > x3)
                    {
                        x3 = Convert.ToInt32(dt.Rows[dt.Rows.Count - 2][j + 6]);
                    }
                    if (Convert.ToInt32(dt.Rows[dt.Rows.Count - 1][j + 6]) > x4)
                    {
                        x4 = Convert.ToInt32(dt.Rows[dt.Rows.Count - 1][j + 6]);
                    }
                }
                worksheet.Cells[dt.Rows.Count, dateTimepd_res.Count + 7].Value = x3;
                worksheet.Cells[dt.Rows.Count + 1, dateTimepd_res.Count + 7].Value = x4;
                for (int j = 0; j < dateTimepd_res.Count; j++) // Выделение цветом максимальных значений
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (Convert.ToInt32(dt.Rows[dt.Rows.Count - 2][j + 6]) == x3)
                    {
                        worksheet.Cells[dt.Rows.Count, j + 7].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[dt.Rows.Count, j + 7].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                    }
                    if (Convert.ToInt32(dt.Rows[dt.Rows.Count - 1][j + 6]) == x4)
                    {
                        worksheet.Cells[dt.Rows.Count + 1, j + 7].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheet.Cells[dt.Rows.Count + 1, j + 7].Style.Fill.BackgroundColor.SetColor(Color.LightGreen);
                    }
                }
                package.Save();
            }
        }
    }
}
