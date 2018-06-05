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
        private static Regex _licenceRegex;
        private static bool _autoCommand;
        private static readonly Dictionary<int, List<string>> Commands = new Dictionary<int, List<string>>();
        private static int _i;
        private static int _j;
        private static readonly Dictionary<string, CarInfo> Cars = new Dictionary<string, CarInfo>();
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
            LoadCommands("Commands.txt");
            _annuledRegex = LoadRegexData("AnnulledRegex.txt");
            _validRegex = LoadRegexData("ValidRegex.txt");
            _licenceRegex = LoadRegexData("LicenceRegex.txt");

            while (true)
            {
                Console.WriteLine("Загрузка данных");
                var directoryPath = ReadInput<string>("Укажите путь к папке с файлами, которые нужно смерджить");

                if (!Directory.Exists(directoryPath))
                {
                    Console.WriteLine($"Папки {directoryPath} не существует");
                    continue;
                }

                var files = GetExcelFiles(directoryPath);

                foreach (var file in files)
                {
                    ProceedFile(file);

                    Console.WriteLine();
                    Console.WriteLine($"Загрузка файла {file.Name} успешно завершена");
                    Console.WriteLine();
                }

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
            var cars = Cars.Where(t => t.Value.Year >= year && t.Value.Status == CarStatus.Valid).ToList();

            using (var excel = new ExcelPackage(fileName))
            {
                var ws = excel.Workbook.Worksheets.Add("Merged Data");

                ws.Cells["A1"].Value = "ЛИЦЕНЗИЯ";
                ws.Cells["A1"].Style.Font.Bold = true;
                ws.Cells["B1"].Value = "ГОС.НОМЕР";
                ws.Cells["B1"].Style.Font.Bold = true;
                ws.Cells["C1"].Value = "ГОД";
                ws.Cells["C1"].Style.Font.Bold = true;
                ws.Cells["D1"].Value = "МАРКА";
                ws.Cells["D1"].Style.Font.Bold = true;
                ws.Cells["E1"].Value = "МОДЕЛЬ";
                ws.Cells["E1"].Style.Font.Bold = true;

                for (var i = 0; i < cars.Count; i++)
                {
                    var row = i + 2;
                    var car = cars[i].Value;
                    ws.Cells[row, 1].Value = car.Licence;
                    ws.Cells[row, 2].Value = car.RegNumber;
                    ws.Cells[row, 3].Value = car.Year;
                    ws.Cells[row, 4].Value = car.Brend;
                    ws.Cells[row, 5].Value = car.Model;
                }

                excel.Save();
            }

            Console.WriteLine("Готово!");
            Console.ReadLine();
        }

        private static void ProceedFile(FileInfo file)
        {
            Console.WriteLine($"Обработка файла {file.Name}");
            using (var excel = new ExcelPackage(file))
            {
                foreach (var worksheet in excel.Workbook.Worksheets)
                {
                    string proceed;
                    if (_autoCommand)
                    {
                        Console.WriteLine($"Какой набор комманд использовать для вкладки {worksheet.Name}? (1,2,3..) или нет для пропуска");

                        int commands;
                        while (true)
                        {
                            var value = ReadInput<string>();

                            if (!int.TryParse(value, out commands))
                                break;

                            if(commands <= 0)
                            {
                                Console.WriteLine($"Число должно быть положительным, но введено {commands}");
                                continue;
                            }

                            if (commands > Commands.Count)
                            {
                                Console.WriteLine($"Число не может быть больше чем {Commands.Count}");
                                continue;
                            }

                            break;
                        }

                        if(commands == 0)
                            continue;

                        proceed = Commands[commands].Count == 8 ? "ДА1" : "ДА2";
                        _i = commands;
                    }
                    else
                    {
                        Console.WriteLine($"Делать обработку для вкладки с название {worksheet.Name}? (да1/да2/нет)");
                        Console.WriteLine("Если да1 - нужно будет указать регион и столбцы для реестра и бланка");
                        Console.WriteLine("Если да2 - нужно будет указать столбец с уже готовым номером лицензии");
                        proceed = ReadInput<string>();

                        if (string.IsNullOrEmpty(proceed) || (proceed.ToUpper() != "ДА1" && proceed.ToUpper() != "ДА2"))
                            continue;
                    }
                    

                    var cars = new Dictionary<string, CarInfo>();

                    var licenceMerged = false;

                    var region = "";
                    var licenceNumberCol = 0;
                    var blankNumberCol = 0;
                    var licenceMergedCol = 0;

                    Console.WriteLine();
                    if (proceed.ToUpper() == "ДА1")
                    {
                        Console.WriteLine(@"Укажите название региона для этой вкладки (например МО, МСК)");
                        Console.WriteLine(@"Оно так же будет использованно для соединения номера реестра ");
                        Console.WriteLine(@"и номера бланка для получения номера лицензии");
                        Console.WriteLine();
                        region = ReadInput<string>("Название региона:", _autoCommand);
                        Console.WriteLine();
                        Console.WriteLine(@"Так же нужно указать в каких столбцах находятся необходимые данные");
                        Console.WriteLine();
                        Console.WriteLine(@"из этого столбца будут взяты только цифры");
                        licenceNumberCol = ReadInput<int>("номер реестра:", _autoCommand);
                        Console.WriteLine();
                        Console.WriteLine(@"из этого столбца будут взяты только цифры");
                        blankNumberCol = ReadInput<int>("номер бланка:", _autoCommand);
                    }
                    else
                    {
                        licenceMerged = true;
                        Console.WriteLine(@"из этого столбца будут взяты только цифры и буквы и он будет переведен в верхний регистр");
                        licenceMergedCol = ReadInput<int>("номер лицензии:", _autoCommand);
                    }
                    
                    Console.WriteLine();
                    Console.WriteLine(@"из этого столбца будут взяты только цифры и буквы и он будет переведен в верхний регистр");
                    var regNumberCol = ReadInput<int>("гос номер: ", _autoCommand);
                    Console.WriteLine();
                    Console.WriteLine("из этого столбца будут взяты только цифры, если пусто, тогда 0");
                    var yearCol = ReadInput<int>("год выпуска:", _autoCommand);
                    Console.WriteLine();
                    var statusCol = 0;
                    if (!licenceMerged)
                    {
                        Console.WriteLine("данные из этого столбца будут преобразованы в ДЕЙСТВУЮЩЕЕ или АННУЛИРОВАНО");
                        Console.WriteLine("в соответствии с регулярными выражениями заданными в AnnulledRegex.txt и ValidRegex.txt (регистр игнориуется)");
                        Console.WriteLine("если значение будет пустым тогда по умолчанию ДЕЙСТВУЮЩЕЕ");
                        Console.WriteLine("а если соотвествий не найдено тогда останется как есть");
                        statusCol = ReadInput<int>("статус лицензии:", _autoCommand);
                    }
                    Console.WriteLine();
                    Console.WriteLine("данные из этого столбца будут заменены на соотвествующее");
                    Console.WriteLine("при совпадении по ключу из таблицы BreadReplace.txt, иначе останется как есть");
                    var brendCol = ReadInput<int>("название бренда:", _autoCommand);
                    Console.WriteLine();
                    Console.WriteLine("данные из этого столбца будут заменены на соотвествующее");
                    Console.WriteLine("при совпадении по ключу из таблицы ModelReplace.txt, иначе останется как есть");
                    var modelCol = ReadInput<int>("название модели:", _autoCommand);
                    Console.WriteLine();

                    if(_autoCommand)
                        _j = 0;

                    var rows = worksheet.Dimension.Rows;

                    try
                    {
                        for (var i = 1; i <= rows; i++)
                        {
                            string licence;

                            if (licenceMerged)
                            {
                                licence = OnlyNumbersAndLetters(worksheet.Cells[i, licenceMergedCol].Text).ToUpper();

                                if (string.IsNullOrWhiteSpace(licence))
                                    continue;
                            }
                            else
                            {
                                var part1 = OnlyNumbers(worksheet.Cells[i, licenceNumberCol].Text);

                                if (string.IsNullOrWhiteSpace(part1))
                                    continue;

                                var part2 = OnlyNumbers(worksheet.Cells[i, blankNumberCol].Text);
                                licence = $"{part1}{region}{part2}";
                            }

                            if (!_licenceRegex.IsMatch(licence))
                            {
                                Console.WriteLine($"Лицензия {licence} имеет некорректный формат. Будет пропущена");
                                continue;
                            }
                            
                            var regNumber = EngCharReplaceOnRus(OnlyNumbersAndLetters(worksheet.Cells[i, regNumberCol].Text)).ToUpper();
                            var yearStr = OnlyNumbers(worksheet.Cells[i, yearCol].Text);
                            var year = string.IsNullOrWhiteSpace(yearStr) ? 0 : int.Parse(yearStr);
                            var status = licenceMerged ? CarStatus.Valid : ReplaceStatus(worksheet.Cells[i, statusCol].Text);
                            var brend = ReplaceIfFound(BrandsForReplace, worksheet.Cells[i, brendCol].Text);
                            var model = ReplaceIfFound(ModelsForReplace, worksheet.Cells[i, modelCol].Text);

                            var car = new CarInfo
                            {
                                Licence = licence,
                                RegNumber = regNumber,
                                Year = year,
                                Status = status,
                                Brend = brend,
                                Model = model
                            };

                            if (cars.ContainsKey(licence))
                            {
                                Console.WriteLine($"Машина с лицензией {licence} уже была добавлена.");
                                continue;
                            }

                            cars.Add(licence, car);
                        }

                        foreach (var car in cars)
                        {
                            if (Cars.ContainsKey(car.Key))
                            {
                                Console.WriteLine($"Машина с лицензией {car.Key} уже была добавлена.");
                                continue;
                            }

                            Cars.Add(car.Key, car.Value);
                        }
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

        private static T ReadInput<T>(string prompt = null, bool autoCommands = false)
        {
            var validInput = false;
            var result = default(T);
            if(!string.IsNullOrWhiteSpace(prompt)) Console.WriteLine(prompt);
            while (!validInput)
            {
                try
                {
                    string value;
                    if (autoCommands)
                    {
                        value = Commands[_i][_j];
                        _j++;
                        Console.WriteLine(value);
                    }
                    else
                        value = Console.ReadLine();

                    result = (T)Convert.ChangeType(value, typeof(T));
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

        private static CarStatus ReplaceStatus(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return CarStatus.Valid;

            if (_annuledRegex.IsMatch(value))
                return CarStatus.Annuled;

            if (_validRegex.IsMatch(value))
                return CarStatus.Valid;

            Console.WriteLine($"Значение статуса {value} не подходит под правила регулярного выражения.");

            return CarStatus.Undefined;
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

        private static void LoadCommands(string fileName)
        {
            if (!File.Exists(fileName))
                return;

            _autoCommand = true;
            var lines = File.ReadAllLines(fileName);

            for (var i = 1; i <= lines.Length; i++)
            {
                var line = lines[i-1];
                var values = line.Split('#')[0].Split(',').ToList();
                Commands.Add(i, values);
            }
        }

        private static List<FileInfo> GetExcelFiles(string dirPath)
        {
            var dir = Directory.GetFiles(dirPath, "*.xlsx");
            return dir.Select(t => new FileInfo(t)).ToList();
        }
    }

    public class CarInfo
    {
        public string Licence { get; set; }
        public string RegNumber { get; set; }
        public int Year { get; set; }
        public CarStatus Status { get; set; }
        public string Brend { get; set; }
        public string Model { get; set; }
    }

    public enum CarStatus
    {
        Undefined,
        Annuled,
        Valid
    }
}
