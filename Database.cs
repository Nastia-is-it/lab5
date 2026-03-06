using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using Microsoft.VisualBasic;

namespace Laba5
{
    // Класс Страна
    public class Country
    {
        public int Id { get; set; }
        public string Name { get; set; }

        public Country() { }

        public Country(int id, string name)
        {
            Id = id;
            Name = name;
        }

        public override string ToString()
        {
            return $"ID: {Id}, Название: {Name}";
        }
    }

    // Класс Клуб
    public class Club
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int CountryId { get; set; }

        public Club() { }

        public Club(int id, string name, int countryId)
        {
            Id = id;
            Name = name;
            CountryId = countryId;
        }

        public override string ToString()
        {
            return $"ID: {Id}, Название: {Name}, ID страны: {CountryId}";
        }
    }

    // Класс Достижения
    public class Achievement
    {
        public int ClubId { get; set; }
        public int Gold { get; set; }              // Б - Золотые медали
        public int Silver { get; set; }            // В - Серебряные медали
        public int Bronze { get; set; }            // К - Бронзовые медали
        public int NationalCupWins { get; set; }   // ФК - Победы в нац. кубке
        public int NationalCupLosses { get; set; } // Ф - Поражения в финале нац. кубка
        public int ChampionsLeagueWins { get; set; } // ЛЧ
        public int ChampionsLeagueLosses { get; set; } // ФЛЧ
        public int EuropaLeagueWins { get; set; } // ЛЕ
        public int EuropaLeagueLosses { get; set; } // ФЛЕ
        public int CupWinnersCupWins { get; set; } // КОК
        public int CupWinnersCupLosses { get; set; } // ФКОК
        public int ConferenceLeagueWins { get; set; } // ЛК
        public int ConferenceLeagueLosses { get; set; } // ФЛК

        public Achievement() { }

        public Achievement(int clubId, int gold, int silver, int bronze,
                          int nationalCupWins, int nationalCupLosses,
                          int championsLeagueWins, int championsLeagueLosses,
                          int europaLeagueWins, int europaLeagueLosses,
                          int cupWinnersCupWins, int cupWinnersCupLosses,
                          int conferenceLeagueWins, int conferenceLeagueLosses)
        {
            ClubId = clubId;
            Gold = gold;
            Silver = silver;
            Bronze = bronze;
            NationalCupWins = nationalCupWins;
            NationalCupLosses = nationalCupLosses;
            ChampionsLeagueWins = championsLeagueWins;
            ChampionsLeagueLosses = championsLeagueLosses;
            EuropaLeagueWins = europaLeagueWins;
            EuropaLeagueLosses = europaLeagueLosses;
            CupWinnersCupWins = cupWinnersCupWins;
            CupWinnersCupLosses = cupWinnersCupLosses;
            ConferenceLeagueWins = conferenceLeagueWins;
            ConferenceLeagueLosses = conferenceLeagueLosses;
        }

        // Общее количество трофеев
        public int TotalTrophies()
        {
            return Gold + Silver + Bronze + NationalCupWins +
                   ChampionsLeagueWins + EuropaLeagueWins +
                   CupWinnersCupWins + ConferenceLeagueWins;
        }

        public override string ToString()
        {
            return $"Клуб ID: {ClubId}, Золото: {Gold}, Серебро: {Silver}, Бронза: {Bronze}, " +
                   $"Нац. кубки: {NationalCupWins}/{NationalCupLosses}, " +
                   $"ЛЧ: {ChampionsLeagueWins}/{ChampionsLeagueLosses}, " +
                   $"ЛЕ: {EuropaLeagueWins}/{EuropaLeagueLosses}, " +
                   $"КОК: {CupWinnersCupWins}/{CupWinnersCupLosses}, " +
                   $"ЛК: {ConferenceLeagueWins}/{ConferenceLeagueLosses}";
        }
    }

    // Вспомогательный класс для работы с базой данных
    public class DatabaseHelper
    {
        // Коллекции для хранения данных
        private readonly List<Country> countries = new List<Country>();
        private readonly List<Club> clubs = new List<Club>();
        private readonly List<Achievement> achievements = new List<Achievement>();

        // Путь к файлу
        private const string FilePath = "LR5-var5.xls";

        // 1. Загрузка данных из Excel файла
        public void LoadDataFromExcel()
        {
            try
            {
                if (!File.Exists(FilePath))
                {
                    Console.WriteLine($"Файл {FilePath} не найден!");
                    Console.WriteLine("Убедитесь, что файл находится в папке с программой.");
                    return;
                }

                Console.WriteLine($"Загружаем файл: {FilePath}");

                Workbook workbook = new Workbook(FilePath);

                // Очищаем предыдущие данные
                countries.Clear();
                clubs.Clear();
                achievements.Clear();

                // Загружаем данные с каждого листа
                LoadCountries(workbook);
                LoadClubs(workbook);
                LoadAchievements(workbook);

                Console.WriteLine("\nДанные успешно загружены!");
                Console.WriteLine($"Стран: {countries.Count}");
                Console.WriteLine($"Клубов: {clubs.Count}");
                Console.WriteLine($"Достижений: {achievements.Count}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при чтении файла: {ex.Message}");
                Console.WriteLine("Убедитесь, что:");
                Console.WriteLine("1. Файл LR5-vag5.xls находится в папке с программой");
                Console.WriteLine("2. Это файл Excel 97-2003 (.xls)");
                Console.WriteLine("3. Файл не открыт в другой программе");
            }
        }

        // Загрузка стран из листа Excel
        private void LoadCountries(Workbook workbook)
        {
            try
            {
                Worksheet countriesSheet = workbook.Worksheets[0];
                Console.WriteLine($"Читаем лист: {countriesSheet.Name}");

                int rowCount = countriesSheet.Cells.Rows.Count;
                Console.WriteLine($"Всего строк в листе: {rowCount}");

                int loadedCount = 0;

                for (int row = 1; row < rowCount; row++)
                {
                    try
                    {
                        var idCell = countriesSheet.Cells[row, 0];
                        if (idCell.Value == null)
                            continue;

                        int id = 0;
                        if (idCell.Type == CellValueType.IsNumeric)
                        {
                            id = (int)idCell.DoubleValue;
                        }
                        else
                        {
                            string strValue = idCell.StringValue;
                            if (!string.IsNullOrEmpty(strValue))
                            {
                                int.TryParse(strValue, out id);
                            }
                        }

                        if (id == 0)
                            continue;

                        var nameCell = countriesSheet.Cells[row, 1];
                        string name = nameCell.StringValue?.Trim() ?? string.Empty;

                        if (!string.IsNullOrEmpty(name))
                        {
                            countries.Add(new Country { Id = id, Name = name });
                            loadedCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Ошибка в строке {row + 1}: {ex.Message}");
                    }
                }

                Console.WriteLine($"Загружено стран: {loadedCount}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при загрузке стран: {ex.Message}");
            }
        }


        // Загрузка клубов из листа Excel
        private void LoadClubs(Workbook workbook)
        {
            try
            {
                Worksheet clubsSheet = workbook.Worksheets[1];
                Console.WriteLine($"Читаем лист: {clubsSheet.Name}");

                int row = 1;
                while (row <= 500 && clubsSheet.Cells[row, 0].Value != null)
                {
                    try
                    {
                        if (clubsSheet.Cells[row, 0].Value != null)
                        {
                            clubs.Add(new Club
                            {
                                Id = Convert.ToInt32(clubsSheet.Cells[row, 0].DoubleValue, CultureInfo.InvariantCulture),
                                Name = clubsSheet.Cells[row, 1].Value?.ToString() ?? string.Empty,
                                CountryId = Convert.ToInt32(clubsSheet.Cells[row, 2].DoubleValue, CultureInfo.InvariantCulture)
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Ошибка чтения строки {row + 1} в таблице Клубы: {ex.Message}");
                    }
                    row++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при загрузке клубов: {ex.Message}");
            }
        }

        // Загрузка достижений из листа Excel
        private void LoadAchievements(Workbook workbook)
        {
            try
            {
                Worksheet achievementsSheet = workbook.Worksheets[2];
                Console.WriteLine($"Читаем лист: {achievementsSheet.Name}");

                int row = 1;
                while (row <= 1000 && achievementsSheet.Cells[row, 0].Value != null)
                {
                    try
                    {
                        if (achievementsSheet.Cells[row, 0].Value != null)
                        {
                            achievements.Add(new Achievement
                            {
                                ClubId = Convert.ToInt32(achievementsSheet.Cells[row, 0].DoubleValue, CultureInfo.InvariantCulture),
                                Gold = Convert.ToInt32(achievementsSheet.Cells[row, 1].DoubleValue, CultureInfo.InvariantCulture),
                                Silver = Convert.ToInt32(achievementsSheet.Cells[row, 2].DoubleValue, CultureInfo.InvariantCulture),
                                Bronze = Convert.ToInt32(achievementsSheet.Cells[row, 3].DoubleValue, CultureInfo.InvariantCulture),
                                NationalCupWins = Convert.ToInt32(achievementsSheet.Cells[row, 4].DoubleValue, CultureInfo.InvariantCulture),
                                NationalCupLosses = Convert.ToInt32(achievementsSheet.Cells[row, 5].DoubleValue, CultureInfo.InvariantCulture),
                                ChampionsLeagueWins = Convert.ToInt32(achievementsSheet.Cells[row, 6].DoubleValue, CultureInfo.InvariantCulture),
                                ChampionsLeagueLosses = Convert.ToInt32(achievementsSheet.Cells[row, 7].DoubleValue, CultureInfo.InvariantCulture),
                                EuropaLeagueWins = Convert.ToInt32(achievementsSheet.Cells[row, 8].DoubleValue, CultureInfo.InvariantCulture),
                                EuropaLeagueLosses = Convert.ToInt32(achievementsSheet.Cells[row, 9].DoubleValue, CultureInfo.InvariantCulture),
                                CupWinnersCupWins = Convert.ToInt32(achievementsSheet.Cells[row, 10].DoubleValue, CultureInfo.InvariantCulture),
                                CupWinnersCupLosses = Convert.ToInt32(achievementsSheet.Cells[row, 11].DoubleValue, CultureInfo.InvariantCulture),
                                ConferenceLeagueWins = Convert.ToInt32(achievementsSheet.Cells[row, 12].DoubleValue, CultureInfo.InvariantCulture),
                                ConferenceLeagueLosses = Convert.ToInt32(achievementsSheet.Cells[row, 13].DoubleValue, CultureInfo.InvariantCulture)
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Ошибка чтения строки {row + 1} в таблице Достижения: {ex.Message}");
                    }
                    row++;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при загрузке достижений: {ex.Message}");
            }
        }

        // 2. Просмотр всех данных
        public void ViewAllData()
        {
            if (!HasData())
            {
                Console.WriteLine("Данные не загружены. Сначала загрузите данные.");
                return;
            }

            DisplayCountries();
            DisplayClubs();
            DisplayAchievements();
        }

        private void DisplayCountries()
        {
            Console.WriteLine("\n=== СТРАНЫ ===");
            if (countries.Count > 0)
            {
                foreach (var country in countries.OrderBy(c => c.Id))
                {
                    Console.WriteLine(country);
                }
                Console.WriteLine($"Всего: {countries.Count} стран");
            }
            else
            {
                Console.WriteLine("Нет данных");
            }
        }

        private void DisplayClubs()
        {
            Console.WriteLine("\n=== КЛУБЫ ===");
            if (clubs.Count > 0)
            {
                foreach (var club in clubs.OrderBy(c => c.Id))
                {
                    Console.WriteLine(club);
                }
                Console.WriteLine($"Всего: {clubs.Count} клубов");
            }
            else
            {
                Console.WriteLine("Нет данных");
            }
        }

        private void DisplayAchievements()
        {
            Console.WriteLine("\n=== ДОСТИЖЕНИЯ ===");
            if (achievements.Count > 0)
            {
                foreach (var achievement in achievements.OrderBy(a => a.ClubId))
                {
                    Console.WriteLine(achievement);
                }
                Console.WriteLine($"Всего: {achievements.Count} записей о достижениях");
            }
            else
            {
                Console.WriteLine("Нет данных");
            }
        }

        // 3. Удаление элементов по ключу
        public void DeleteCountry(int id)
        {
            // Проверяем, есть ли клубы в этой стране
            if (clubs.Any(c => c.CountryId == id))
            {
                Console.WriteLine($"Нельзя удалить страну с ID {id}, так как в ней есть клубы!");
                return;
            }

            var country = countries.FirstOrDefault(c => c.Id == id);
            if (country != null)
            {
                countries.Remove(country);
                Console.WriteLine($"Страна с ID {id} удалена.");
            }
            else
            {
                Console.WriteLine($"Страна с ID {id} не найдена.");
            }
        }

        public void DeleteClub(int id)
        {
            // Проверяем, есть ли достижения у этого клуба
            if (achievements.Any(a => a.ClubId == id))
            {
                Console.WriteLine($"Сначала удалите достижения клуба с ID {id}!");
                return;
            }

            var club = clubs.FirstOrDefault(c => c.Id == id);
            if (club != null)
            {
                clubs.Remove(club);
                Console.WriteLine($"Клуб с ID {id} удален.");
            }
            else
            {
                Console.WriteLine($"Клуб с ID {id} не найден.");
            }
        }

        public void DeleteAchievement(int clubId)
        {
            var achievement = achievements.FirstOrDefault(a => a.ClubId == clubId);
            if (achievement != null)
            {
                achievements.Remove(achievement);
                Console.WriteLine($"Достижения для клуба с ID {clubId} удалены.");
            }
            else
            {
                Console.WriteLine($"Достижения для клуба с ID {clubId} не найдены.");
            }
        }

        // 4. Добавление элементов
        public void AddCountry(Country country)
        {
            if (country == null)
                throw new ArgumentNullException(nameof(country));

            if (countries.Any(c => c.Id == country.Id))
            {
                Console.WriteLine($"Страна с ID {country.Id} уже существует!");
                return;
            }

            countries.Add(country);
            Console.WriteLine("Страна добавлена.");
        }

        public void AddClub(Club club)
        {
            if (club == null)
                throw new ArgumentNullException(nameof(club));

            if (clubs.Any(c => c.Id == club.Id))
            {
                Console.WriteLine($"Клуб с ID {club.Id} уже существует!");
                return;
            }

            if (!countries.Any(c => c.Id == club.CountryId))
            {
                Console.WriteLine($"Страны с ID {club.CountryId} не существует!");
                return;
            }

            clubs.Add(club);
            Console.WriteLine("Клуб добавлен.");
        }

        public void AddAchievement(Achievement achievement)
        {
            if (achievement == null)
                throw new ArgumentNullException(nameof(achievement));

            if (achievements.Any(a => a.ClubId == achievement.ClubId))
            {
                Console.WriteLine($"Достижения для клуба с ID {achievement.ClubId} уже существуют!");
                return;
            }

            if (!clubs.Any(c => c.Id == achievement.ClubId))
            {
                Console.WriteLine($"Клуба с ID {achievement.ClubId} не существует!");
                return;
            }

            achievements.Add(achievement);
            Console.WriteLine("Достижения добавлены.");
        }

        // 5. ЗАПРОСЫ

        // 5.1 Запрос к одной таблице (возвращает перечень)
        // Клубы, выигравшие Лигу Чемпионов
        public void GetChampionsLeagueWinners()
        {
            var clubIds = achievements
                .Where(a => a.ChampionsLeagueWins > 0)
                .Select(a => a.ClubId)
                .ToList();

            var result = clubs
                .Where(c => clubIds.Contains(c.Id))
                .OrderBy(c => c.Name)
                .ToList();

            Console.WriteLine("\n=== КЛУБЫ, ВЫИГРАВШИЕ ЛИГУ ЧЕМПИОНОВ ===");
            if (result.Count > 0)
            {
                foreach (var club in result)
                {
                    var country = countries.FirstOrDefault(c => c.Id == club.CountryId);
                    Console.WriteLine($"{club.Name} ({country?.Name ?? "Неизвестно"})");
                }
                Console.WriteLine($"Всего: {result.Count} клубов");
            }
            else
            {
                Console.WriteLine("Нет таких клубов.");
            }
        }

        // 5.2 Запрос к двум таблицам (возвращает одно значение)
        // Общее количество трофеев для клуба по названию
        public void GetTotalTrophiesByClubName(string clubName)
        {
            var club = clubs.FirstOrDefault(c =>
                c.Name.IndexOf(clubName, StringComparison.OrdinalIgnoreCase) >= 0);

            if (club == null)
            {
                Console.WriteLine($"Клуб '{clubName}' не найден.");
                return;
            }

            var achievement = achievements.FirstOrDefault(a => a.ClubId == club.Id);
            int total = achievement?.TotalTrophies() ?? 0;

            Console.WriteLine($"\nКлуб: {club.Name}");
            Console.WriteLine($"Общее количество трофеев: {total}");
        }

        // 5.3 Запросы к трем таблицам

        // Запрос 1: возвращает перечень - клубы с их странами и общим количеством трофеев
        public void GetClubsWithCountriesAndTotalTrophies()
        {
            var query = from club in clubs
                        join country in countries on club.CountryId equals country.Id
                        join ach in achievements on club.Id equals ach.ClubId into achievementsGroup
                        from ach in achievementsGroup.DefaultIfEmpty()
                        orderby (ach != null ? ach.TotalTrophies() : 0) descending
                        select new
                        {
                            ClubName = club.Name,
                            Country = country.Name,
                            TotalTrophies = ach != null ? ach.TotalTrophies() : 0,
                            EuropeanTrophies = ach != null ?
                                ach.ChampionsLeagueWins + ach.EuropaLeagueWins +
                                ach.CupWinnersCupWins + ach.ConferenceLeagueWins : 0
                        };

            Console.WriteLine("\n=== КЛУБЫ С ИХ СТРАНАМИ И КОЛИЧЕСТВОМ ТРОФЕЕВ ===");
            int count = 0;
            foreach (var item in query)
            {
                Console.WriteLine($"{item.ClubName} ({item.Country})");
                Console.WriteLine($"  Всего трофеев: {item.TotalTrophies}");
                Console.WriteLine($"  Еврокубки: {item.EuropeanTrophies}");
                Console.WriteLine();
                count++;
                if (count >= 20) // Ограничим вывод для читаемости
                {
                    Console.WriteLine($"... и еще {query.Count() - count} записей");
                    break;
                }
            }
        }

        // Запрос 2: возвращает одно значение - страна с наибольшим количеством трофеев
        public void GetCountryWithMostTrophies()
        {
            var query = from club in clubs
                        join ach in achievements on club.Id equals ach.ClubId
                        join country in countries on club.CountryId equals country.Id
                        group new { ach, country } by country.Name into g
                        select new
                        {
                            CountryName = g.Key,
                            TotalTrophies = g.Sum(x => x.ach.TotalTrophies()),
                            TotalClubs = g.Select(x => x.ach.ClubId).Distinct().Count()
                        };

            var result = query.OrderByDescending(x => x.TotalTrophies).FirstOrDefault();

            Console.WriteLine("\n=== СТРАНА С НАИБОЛЬШИМ КОЛИЧЕСТВОМ ТРОФЕЕВ ===");
            if (result != null)
            {
                Console.WriteLine($"Страна: {result.CountryName}");
                Console.WriteLine($"Всего трофеев: {result.TotalTrophies}");
                Console.WriteLine($"Количество клубов: {result.TotalClubs}");
            }
            else
            {
                Console.WriteLine("Нет данных");
            }
        }

        // Запрос 3: возвращает перечень - клубы, выигравшие еврокубки
        public void GetClubsWithEuropeanTrophies()
        {
            var query = from club in clubs
                        join ach in achievements on club.Id equals ach.ClubId
                        where ach.ChampionsLeagueWins > 0 || ach.EuropaLeagueWins > 0
                              || ach.CupWinnersCupWins > 0 || ach.ConferenceLeagueWins > 0
                        join country in countries on club.CountryId equals country.Id
                        orderby country.Name, club.Name
                        select new
                        {
                            ClubName = club.Name,
                            Country = country.Name,
                            CL = ach.ChampionsLeagueWins,
                            EL = ach.EuropaLeagueWins,
                            CWC = ach.CupWinnersCupWins,
                            Conf = ach.ConferenceLeagueWins,
                            Total = ach.ChampionsLeagueWins + ach.EuropaLeagueWins +
                                    ach.CupWinnersCupWins + ach.ConferenceLeagueWins
                        };

            Console.WriteLine("\n=== КЛУБЫ, ВЫИГРАВШИЕ ЕВРОКУБКИ ===");
            int count = 0;
            foreach (var item in query)
            {
                Console.WriteLine($"{item.ClubName} ({item.Country}) - всего: {item.Total}");
                Console.WriteLine($"  ЛЧ: {item.CL}, ЛЕ: {item.EL}, КОК: {item.CWC}, ЛК: {item.Conf}");
                Console.WriteLine();
                count++;
                if (count >= 15)
                {
                    Console.WriteLine($"... и еще {query.Count() - count} записей");
                    break;
                }
            }
        }

        // Запрос 4: возвращает одно значение - клуб с наибольшим количеством серебряных медалей
        public void GetClubWithMostSilverMedals()
        {
            var query = from ach in achievements
                        join club in clubs on ach.ClubId equals club.Id
                        join country in countries on club.CountryId equals country.Id
                        orderby ach.Silver descending
                        select new
                        {
                            ClubName = club.Name,
                            Country = country.Name,
                            Silver = ach.Silver,
                            Gold = ach.Gold,
                            Bronze = ach.Bronze
                        };

            var result = query.FirstOrDefault();

            Console.WriteLine("\n=== КЛУБ С НАИБОЛЬШИМ КОЛИЧЕСТВОМ СЕРЕБРЯНЫХ МЕДАЛЕЙ ===");
            if (result != null)
            {
                Console.WriteLine($"Клуб: {result.ClubName} ({result.Country})");
                Console.WriteLine($"Серебряных медалей: {result.Silver}");
                Console.WriteLine($"Золотых: {result.Gold}, Бронзовых: {result.Bronze}");
            }
            else
            {
                Console.WriteLine("Нет данных");
            }
        }

        // 6. Сохранение данных в Excel файл
        public void SaveToExcel()
        {
            try
            {
                const string outputFile = "LR5-vag5_modified.xls";

                Workbook workbook = new Workbook();

                SaveCountriesToWorksheet(workbook);
                SaveClubsToWorksheet(workbook);
                SaveAchievementsToWorksheet(workbook);

                workbook.Save(outputFile, SaveFormat.Excel97To2003);

                Console.WriteLine($"\nДанные успешно сохранены в файл: {outputFile}");
                Console.WriteLine($"Расположение файла: {Path.GetFullPath(outputFile)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при сохранении: {ex.Message}");
            }
        }

        private void SaveCountriesToWorksheet(Workbook workbook)
        {
            Worksheet countriesSheet = workbook.Worksheets[0];
            countriesSheet.Name = "Страна";

            // Заголовки
            countriesSheet.Cells["A1"].PutValue("ID страны");
            countriesSheet.Cells["B1"].PutValue("Название");

            // Данные
            for (int i = 0; i < countries.Count; i++)
            {
                countriesSheet.Cells[i + 1, 0].PutValue(countries[i].Id);
                countriesSheet.Cells[i + 1, 1].PutValue(countries[i].Name);
            }

            countriesSheet.AutoFitColumns();
        }

        private void SaveClubsToWorksheet(Workbook workbook)
        {
            Worksheet clubsSheet = workbook.Worksheets.Add("Клубы");

            // Заголовки
            clubsSheet.Cells["A1"].PutValue("ID клуба");
            clubsSheet.Cells["B1"].PutValue("Название");
            clubsSheet.Cells["C1"].PutValue("ID страны");

            // Данные
            for (int i = 0; i < clubs.Count; i++)
            {
                clubsSheet.Cells[i + 1, 0].PutValue(clubs[i].Id);
                clubsSheet.Cells[i + 1, 1].PutValue(clubs[i].Name);
                clubsSheet.Cells[i + 1, 2].PutValue(clubs[i].CountryId);
            }

            clubsSheet.AutoFitColumns();
        }

        private void SaveAchievementsToWorksheet(Workbook workbook)
        {
            Worksheet achievementsSheet = workbook.Worksheets.Add("Достижения");

            // Заголовки
            achievementsSheet.Cells["A1"].PutValue("ID клуба");
            achievementsSheet.Cells["B1"].PutValue("Б");
            achievementsSheet.Cells["C1"].PutValue("В");
            achievementsSheet.Cells["D1"].PutValue("К");
            achievementsSheet.Cells["E1"].PutValue("ФК");
            achievementsSheet.Cells["F1"].PutValue("Ф");
            achievementsSheet.Cells["G1"].PutValue("ЛЧ");
            achievementsSheet.Cells["H1"].PutValue("ФЛЧ");
            achievementsSheet.Cells["I1"].PutValue("ЛЕ");
            achievementsSheet.Cells["J1"].PutValue("ФЛЕ");
            achievementsSheet.Cells["K1"].PutValue("КОК");
            achievementsSheet.Cells["L1"].PutValue("ФКОК");
            achievementsSheet.Cells["M1"].PutValue("ЛК");
            achievementsSheet.Cells["N1"].PutValue("ФЛК");

            // Данные
            for (int i = 0; i < achievements.Count; i++)
            {
                var a = achievements[i];
                achievementsSheet.Cells[i + 1, 0].PutValue(a.ClubId);
                achievementsSheet.Cells[i + 1, 1].PutValue(a.Gold);
                achievementsSheet.Cells[i + 1, 2].PutValue(a.Silver);
                achievementsSheet.Cells[i + 1, 3].PutValue(a.Bronze);
                achievementsSheet.Cells[i + 1, 4].PutValue(a.NationalCupWins);
                achievementsSheet.Cells[i + 1, 5].PutValue(a.NationalCupLosses);
                achievementsSheet.Cells[i + 1, 6].PutValue(a.ChampionsLeagueWins);
                achievementsSheet.Cells[i + 1, 7].PutValue(a.ChampionsLeagueLosses);
                achievementsSheet.Cells[i + 1, 8].PutValue(a.EuropaLeagueWins);
                achievementsSheet.Cells[i + 1, 9].PutValue(a.EuropaLeagueLosses);
                achievementsSheet.Cells[i + 1, 10].PutValue(a.CupWinnersCupWins);
                achievementsSheet.Cells[i + 1, 11].PutValue(a.CupWinnersCupLosses);
                achievementsSheet.Cells[i + 1, 12].PutValue(a.ConferenceLeagueWins);
                achievementsSheet.Cells[i + 1, 13].PutValue(a.ConferenceLeagueLosses);
            }

            achievementsSheet.AutoFitColumns();
        }

        // Проверка наличия данных
        public bool HasData()
        {
            return countries.Count > 0 || clubs.Count > 0 || achievements.Count > 0;
        }

        // Вспомогательные методы для проверки существования
        public bool CountryExists(int id)
        {
            return countries.Any(c => c.Id == id);
        }

        public bool ClubExists(int id)
        {
            return clubs.Any(c => c.Id == id);
        }

        public void ListCountries()
        {
            Console.WriteLine("\nДоступные страны:");
            foreach (var country in countries.OrderBy(c => c.Id))
            {
                Console.WriteLine($"  {country.Id}: {country.Name}");
            }
        }

        public void ListClubs()
        {
            Console.WriteLine("\nДоступные клубы:");
            foreach (var club in clubs.OrderBy(c => c.Id).Take(20))
            {
                var country = countries.FirstOrDefault(c => c.Id == club.CountryId);
                Console.WriteLine($"  {club.Id}: {club.Name} ({country?.Name ?? "Неизвестно"})");
            }
            if (clubs.Count > 20)
                Console.WriteLine($"  ... и еще {clubs.Count - 20} клубов");
        }
    }
}
