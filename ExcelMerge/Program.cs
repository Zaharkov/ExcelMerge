using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace ExcelMerge
{
    class Program
    {
        private static Regex _annuledRegex;
        private static Regex _validRegex;
        private static readonly List<CarInfo> Cars = new List<CarInfo>();
        private static readonly Dictionary<string, string> ModelsForReplace = new Dictionary<string, string>();
        private static readonly Dictionary<string, string> BrandsForReplace = new Dictionary<string, string>();
        private static readonly KeyValuePair<string, string> CharReplace = new KeyValuePair<string, string>
        (
            "thbmacekopxyACEHKMOPTXBY", 
            "тнвмасекорхуАСЕНКМОРТХВУ"
        );

        public static void Main(string[] args)
        {
            LoadReplaceData(BrandsForReplace, "BrandReplace.txt");
            LoadReplaceData(ModelsForReplace, "ModelReplace.txt");
            _annuledRegex = LoadRegexData("AnnulledRegex.txt");
            _validRegex = LoadRegexData("ValidRegex.txt");

            while (true)
            {
                Console.WriteLine("Загрузка данных");
                var filePath = ReadInput<string>("Укажите путь к файлу вместе с название файла");

                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Файла {filePath} не существует");
                    continue;
                }

                ProceedFile(new FileInfo(filePath));

                Console.WriteLine();
                Console.WriteLine($"Загрузка файла {filePath} успешно завершена");
                Console.WriteLine();
                var proceed = ReadInput<string>("Хотите загрузить ещё данные? (да/нет)");
                Console.WriteLine();

                if (string.IsNullOrEmpty(proceed) || proceed.ToUpper() != "ДА")
                    break;
            }

            Console.WriteLine("Создаем таблицу на основен данных");

            FileInfo fileName;
            while (true)
            {
                var name = ReadInput<string>("Введите название файла");
                fileName = new FileInfo($"{name}.xlsx");
                if (fileName.Exists)
                {
                    Console.WriteLine($"Файл {fileName.DirectoryName}/{fileName.Name} уже существует, введите другое имя");
                    continue;
                }

                break;
            }

            var year = ReadInput<int>("Начиная с какого года (включая) делать выборку?");
            var cars = Cars.Where(t => t.Year >= year).ToList();

            using (var excel = new ExcelPackage(fileName))
            {
                var ws = excel.Workbook.Worksheets.Add("Merged Data");

                ws.Cells["A1"].Value = "РЕГИОН";
                ws.Cells["A1"].Style.Font.Bold = true;
                ws.Cells["B1"].Value = "ЛИЦЕНЗИЯ";
                ws.Cells["B1"].Style.Font.Bold = true;
                ws.Cells["C1"].Value = "ГОС.НОМЕР";
                ws.Cells["C1"].Style.Font.Bold = true;
                ws.Cells["D1"].Value = "ГОД";
                ws.Cells["D1"].Style.Font.Bold = true;
                ws.Cells["E1"].Value = "СТАТУС";
                ws.Cells["E1"].Style.Font.Bold = true;
                ws.Cells["F1"].Value = "МАРКА";
                ws.Cells["F1"].Style.Font.Bold = true;
                ws.Cells["G1"].Value = "МОДЕЛЬ";
                ws.Cells["G1"].Style.Font.Bold = true;

                for (var i = 0; i < cars.Count; i++)
                {
                    var row = i + 2;
                    var car = cars[i];
                    ws.Cells[row, 1].Value = car.Region;
                    ws.Cells[row, 2].Value = car.Licence;
                    ws.Cells[row, 3].Value = car.RegNumber;
                    ws.Cells[row, 4].Value = car.Year;
                    ws.Cells[row, 5].Value = car.Status;
                    ws.Cells[row, 6].Value = car.Brend;
                    ws.Cells[row, 7].Value = car.Model;
                }

                excel.Save();
            }

            Console.WriteLine("Готово!");
            Console.ReadLine();
        }

        private static void ProceedFile(FileInfo file)
        {
            using (var excel = new ExcelPackage(file))
            {
                foreach (var worksheet in excel.Workbook.Worksheets)
                {
                    var proceed = ReadInput<string>($"Делать обработку для вкладки с название {worksheet.Name}? (да/нет)");

                    if(string.IsNullOrEmpty(proceed) || proceed.ToUpper() != "ДА")
                        continue;

                    var cars = new List<CarInfo>();
                    Console.WriteLine();
                    Console.WriteLine(@"Укажите название региона для этой вкладки (например МО, МСК)");
                    Console.WriteLine(@"Оно так же будет использованно для соединения номера реестра ");
                    Console.WriteLine(@"и номера бланка для получения номера лицензии");
                    Console.WriteLine();
                    var region = ReadInput<string>("Название региона:");
                    Console.WriteLine();
                    Console.WriteLine(@"Так же нужно указать в каких столбцах находятся необходимые данные");
                    Console.WriteLine();
                    Console.WriteLine(@"из этого столбца будут взяты только цифры");
                    var licenceNumberCol = ReadInput<int>("номер реестра:");
                    Console.WriteLine();
                    Console.WriteLine(@"из этого столбца будут взяты только цифры");
                    var blankNumberCol = ReadInput<int>("номер бланка:");
                    Console.WriteLine();
                    Console.WriteLine(@"из этого столбца будут взяты только цифры и буквы и он будет переведен в верхний регистр");
                    var regNumberCol = ReadInput<int>("гос номер: ");
                    Console.WriteLine();
                    Console.WriteLine("из этого столбца будут взяты только цифры, если пусто, тогда 0");
                    var yearCol = ReadInput<int>("год выпуска:");
                    Console.WriteLine();
                    Console.WriteLine("данные из этого столбца будут преобразованы в ДЕЙСТВУЮЩЕЕ или АННУЛИРОВАНО");
                    Console.WriteLine("в соответствии с регулярными выражениями заданными в AnnulledRegex.txt и ValidRegex.txt (регистр игнориуется)");
                    Console.WriteLine("если значение будет пустым тогда по умолчанию ДЕЙСТВУЮЩЕЕ");
                    Console.WriteLine("а если соотвествий не найдено тогда останется как есть");
                    var statusCol = ReadInput<int>("статус лицензии:");
                    Console.WriteLine();
                    Console.WriteLine("данные из этого столбца будут заменены на соотвествующее");
                    Console.WriteLine("при совпадении по ключу из таблицы BreadReplace.txt, иначе останется как есть");
                    var brendCol = ReadInput<int>("название бренда:");
                    Console.WriteLine();
                    Console.WriteLine("данные из этого столбца будут заменены на соотвествующее");
                    Console.WriteLine("при совпадении по ключу из таблицы ModelReplace.txt, иначе останется как есть");
                    var modelCol = ReadInput<int>("название модели:");
                    Console.WriteLine();

                    var rows = worksheet.Dimension.Rows;

                    try
                    {
                        for (var i = 1; i <= rows; i++)
                        {
                            var part1 = OnlyNumbers(worksheet.Cells[i, licenceNumberCol].Text);

                            if (string.IsNullOrWhiteSpace(part1))
                                continue;

                            var part2 = OnlyNumbers(worksheet.Cells[i, blankNumberCol].Text);
                            var licence = $"{part1}{region}{part2}";

                            var regNumber = EngCharReplaceOnRus(OnlyNumbersAndLetters(worksheet.Cells[i, regNumberCol].Text)).ToUpper();
                            var yearStr = OnlyNumbers(worksheet.Cells[i, yearCol].Text);
                            var year = string.IsNullOrWhiteSpace(yearStr) ? 0 : int.Parse(yearStr);
                            var status = ReplaceStatus(worksheet.Cells[i, statusCol].Text);
                            var brend = ReplaceIfFound(BrandsForReplace, worksheet.Cells[i, brendCol].Text);
                            var model = ReplaceIfFound(ModelsForReplace, worksheet.Cells[i, modelCol].Text);

                            cars.Add(new CarInfo
                            {
                                Region = region,
                                Licence = licence,
                                RegNumber = regNumber,
                                Year = year,
                                Status = status,
                                Brend = brend,
                                Model = model
                            });
                        }

                        Cars.AddRange(cars);
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"Произошла ошибка при обработке вкладки с название {worksheet.Name}");
                        Console.WriteLine("Возможно вы указали некорректные номера столбцов");
                        Console.WriteLine(e.ToString());
                        Console.ReadLine();
                        continue;
                    }

                    Console.WriteLine($"Загруженно дополнительно {cars.Count} информации по машинам");
                    Console.WriteLine($"Текущий общий размер - {Cars.Count}");
                }
            }
        }

        private static T ReadInput<T>(string prompt)
        {
            var validInput = false;
            var result = default(T);
            Console.WriteLine(prompt);
            while (!validInput)
            {
                try
                {
                    result = (T)Convert.ChangeType(Console.ReadLine(), typeof(T));
                    validInput = true;
                }
                catch
                {
                    Console.WriteLine("Введенные данные некорректны - попробуйте ещё раз");
                }
            }
            return result;
        }

        private static string OnlyNumbers(string value)
        {
            return new string(value.Where(char.IsDigit).ToArray());
        }

        private static string OnlyNumbersAndLetters(string value)
        {
            return new string(value.Where(char.IsLetterOrDigit).ToArray());
        }

        private static string ReplaceIfFound(IReadOnlyDictionary<string, string> dic, string value)
        {
            var key = new string(value.Where(c => !char.IsSeparator(c)).ToArray());
            return dic.ContainsKey(key.ToUpper()) ? dic[key.ToUpper()] : key;
        }

        private static string ReplaceStatus(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "ДЕЙСТВУЮЩЕЕ";

            if (_annuledRegex.IsMatch(value))
                return "АННУЛИРОВАНО";

            if (_validRegex.IsMatch(value))
                return "ДЕЙСТВУЮЩЕЕ";

            Console.WriteLine($"Значение статуса {value} не подходит под правила регулярного выражения. Будет оставлено как есть");

            return value;
        }

        private static string EngCharReplaceOnRus(string value)
        {
            var engs = CharReplace.Key;
            var russ = CharReplace.Value;

            for (var i = 0; i < engs.Length; i++)
            {
                var eng = engs[i];
                var rus = russ[i];
                value = value.Replace(eng, rus);
            }

            return value;
        }

        private static void LoadReplaceData(IDictionary<string, string> dic, string fileName)
        {
            var lines = File.ReadAllLines(fileName);

            foreach (var line in lines)
            {
                var split = line.Split(new[] {" = "}, StringSplitOptions.None);
                var key = split[0].ToUpper();
                var value = split[1];

                if (dic.ContainsKey(key))
                    dic[key] = value;
                else 
                    dic.Add(key, value);
            }
        }

        private static Regex LoadRegexData(string fileName)
        {
            return new Regex(File.ReadAllText(fileName), RegexOptions.IgnoreCase);
        }
    }

    public class CarInfo
    {
        public string Region { get; set; }
        public string Licence { get; set; }
        public string RegNumber { get; set; }
        public int Year { get; set; }
        public string Status { get; set; }
        public string Brend { get; set; }
        public string Model { get; set; }
    }
}
