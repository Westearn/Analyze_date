using System;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;
using System.Threading;
using System.Runtime;
 /*
 Данная программа предназначена для обработки исходных данных Заказчика ООО "СевКомНефтегаз"
 Алгоритм прозволяет определить количество и мощность максимального числа одновременно работающих скважин
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
                    for (int k = 1; k <= 24; k++)
                    {
                        var datein = Convert.ToDateTime(sheet.Cells[11 + k, 10].Value.ToString().Trim());
                        if (sheet.Cells[11 + k, 11].Value != null)
                        {
                            var dateout = Convert.ToDateTime(sheet.Cells[11 + k, 11].Value.ToString().Trim());
                            variables.dateTimeOut.Add(dateout);
                        }
                        variables.dateTimeIn.Add(datein);
                    }
                    // variables.dateTimeIn.Add(Convert.ToDateTime(sheet.Cells["J12:J35"].Value.ToString().Trim())); // Сбор списка дат ввода скважин
                    // variables.dateTimeOut.Add(Convert.ToDateTime(sheet.Cells["K12:K35"].Value.ToString().Trim())); // Сбор списка дат перевода скважин
                    for (int j = 1; j <= 24; j++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        dt.Rows.Add(new Object[] { number_kp, j, sheet.Cells[11 + j, 2], sheet.Cells[11 + j, 10].Value.ToString().Trim(),
                        sheet.Cells[11 + j, 11].Value == null ? string.Empty : sheet.Cells[11 + j, 11].Value.ToString().Trim(), Convert.ToInt32(sheet.Cells[11 + j, 12].Value) }); // Заполнение таблицы информацией с каждой КП
                    }
                    // for (int k = 1; k < 24; k++)
                    // {
                        // var datein = Convert.ToDateTime(sheet.Cells[10 + k, 9].Value.ToString().Trim());
                        // var dateout = Convert.ToDateTime(sheet.Cells[10 + k, 9].Value.ToString().Trim());
                        // variables.dateTimeIn.Add(datein);
                        // variables.dateTimeOut.Add(dateout);
                    // }
                    poryad++;
                }
            }
            variables.dateTimepd.Concat(variables.dateTimeIn);
            variables.dateTimepd.Concat(variables.dateTimeOut); // Объединение спсиков дат в общий
            variables.dateTimepd.RemoveAll(item => item == null); // Удаление пустых элементов
            variables.dateTimepd.Sort(); // Сортировка списка по возрастанию
            variables.dateTimepd.Distinct(); // Удаление из списка повторяющихся элементов
            foreach (var item in variables.dateTimepd)
            {
                cancellationToken.ThrowIfCancellationRequested();
                dt.Columns.Add(new DataColumn(item.ToShortDateString(), Type.GetType("System.String")));
            }
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                for (int j = 0; j < variables.dateTimepd.Count; j++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if ((Convert.ToDateTime(dt.Rows[i + 1]["Дата ввода скважины"]) >= variables.dateTimepd[j]) && (!dt.Rows[i + 1]["Дата перевода в ППД"].Equals(null) || variables.dateTimepd[j] <= Convert.ToDateTime(dt.Rows[i + 1]["Дата перевода в ППД"])))
                    {
                        dt.Rows[i + 1][j + 6] = dt.Rows[i + 1]["Мощность"];
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
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];  // Обращаемся к первому листу
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    for (int j = 0; j < variables.dateTimepd.Count + 6; j++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        worksheet.Cells[i, j].Value = dt.Rows[i][j]; // Производим запись таблицы в файл
                        package.Save();
                    }
                }
            }
        }
    }
}
