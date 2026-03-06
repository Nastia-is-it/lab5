using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Laba5
{
    class Program
    {
        static DatabaseHelper dbHelper = new DatabaseHelper();

        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;

            while (true)
            {
                try
                {
                    Console.Clear();
                    Console.WriteLine("ФУТБОЛЬНЫЕ КЛУБЫ");
                    Console.WriteLine("1. Загрузить данные из Excel");
                    Console.WriteLine("2. Просмотреть все данные");
                    Console.WriteLine("3. Удалить элемент");
                    Console.WriteLine("4. Добавить элемент");
                    Console.WriteLine("5. Выполнить запросы");
                    Console.WriteLine("6. Сохранить изменения в Excel");
                    Console.WriteLine("0. Выход");
                    Console.WriteLine();
                    Console.Write("Выберите действие: ");

                    string choice = Console.ReadLine();

                    switch (choice)
                    {
                        case "1":
                            LoadData();
                            break;
                        case "2":
                            if (CheckDataLoaded()) ViewAllData();
                            break;
                        case "3":
                            if (CheckDataLoaded()) DeleteMenu();
                            break;
                        case "4":
                            if (CheckDataLoaded()) AddMenu();
                            break;
                        case "5":
                            if (CheckDataLoaded()) QueryMenu();
                            break;
                        case "6":
                            if (CheckDataLoaded()) SaveData();
                            break;
                        case "0":
                            Console.WriteLine("Программа завершена.");
                            return;
                        default:
                            Console.WriteLine("Неверный выбор! Нажмите любую клавишу...");
                            Console.ReadKey();
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка: {ex.Message}");
                    Console.WriteLine("Нажмите любую клавишу для продолжения...");
                    Console.ReadKey();
                }
            }
        }

        static bool CheckDataLoaded()
        {
            if (!dbHelper.HasData())
            {
                Console.WriteLine("Данные не загружены! Сначала загрузите данные.");
                Console.WriteLine("Нажмите любую клавишу...");
                Console.ReadKey();
                return false;
            }
            return true;
        }

        static void LoadData()
        {
            Console.WriteLine("\nЗагрузка данных");
            dbHelper.LoadDataFromExcel();
            Console.WriteLine("\nНажмите любую клавишу...");
            Console.ReadKey();
        }

        static void ViewAllData()
        {
            Console.WriteLine("\nПросмотр данных");
            dbHelper.ViewAllData();
            Console.WriteLine("\nНажмите любую клавишу...");
            Console.ReadKey();
        }

        static void DeleteMenu()
        {
            Console.WriteLine("\nУдаление элемента");
            Console.WriteLine("1. Удалить страну");
            Console.WriteLine("2. Удалить клуб");
            Console.WriteLine("3. Удалить достижения");
            Console.Write("Выберите: ");

            string choice = Console.ReadLine();

            switch (choice)
            {
                case "1":
                    dbHelper.ListCountries();
                    Console.Write("Введите ID страны для удаления: ");
                    if (int.TryParse(Console.ReadLine(), out int countryId))
                        dbHelper.DeleteCountry(countryId);
                    else
                        Console.WriteLine("Неверный формат ID!");
                    break;

                case "2":
                    dbHelper.ListClubs();
                    Console.Write("Введите ID клуба для удаления: ");
                    if (int.TryParse(Console.ReadLine(), out int clubId))
                        dbHelper.DeleteClub(clubId);
                    else
                        Console.WriteLine("Неверный формат ID!");
                    break;

                case "3":
                    Console.Write("Введите ID клуба для удаления достижений: ");
                    if (int.TryParse(Console.ReadLine(), out int achClubId))
                        dbHelper.DeleteAchievement(achClubId);
                    else
                        Console.WriteLine("Неверный формат ID!");
                    break;

                default:
                    Console.WriteLine("Неверный выбор!");
                    break;
            }

            Console.WriteLine("\nНажмите любую клавишу...");
            Console.ReadKey();
        }

        static void AddMenu()
        {
            Console.WriteLine("\nДобавление элемента");
            Console.WriteLine("1. Добавить страну");
            Console.WriteLine("2. Добавить клуб");
            Console.WriteLine("3. Добавить достижения");
            Console.Write("Выберите: ");

            string choice = Console.ReadLine();

            try
            {
                switch (choice)
                {
                    case "1":
                        AddCountry();
                        break;
                    case "2":
                        AddClub();
                        break;
                    case "3":
                        AddAchievement();
                        break;
                    default:
                        Console.WriteLine("Неверный выбор!");
                        break;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка при добавлении: {ex.Message}");
            }

            Console.WriteLine("\nНажмите любую клавишу...");
            Console.ReadKey();
        }

        static void AddCountry()
        {
            Console.WriteLine("\nДобавление страны");
            Console.Write("ID страны: ");
            int id = int.Parse(Console.ReadLine());

            Console.Write("Название страны: ");
            string name = Console.ReadLine();

            dbHelper.AddCountry(new Country(id, name));
        }

        static void AddClub()
        {
            Console.WriteLine("\nДобавление клуба");
            dbHelper.ListCountries();

            Console.Write("ID клуба: ");
            int id = int.Parse(Console.ReadLine());

            Console.Write("Название клуба: ");
            string name = Console.ReadLine();

            Console.Write("ID страны: ");
            int countryId = int.Parse(Console.ReadLine());

            dbHelper.AddClub(new Club(id, name, countryId));
        }

        static void AddAchievement()
        {
            Console.WriteLine("\nДобавление достижений");
            dbHelper.ListClubs();

            Console.Write("ID клуба: ");
            int clubId = int.Parse(Console.ReadLine());

            Console.WriteLine("Введите количество (числа):");

            Console.Write("Золотые медали (Б): ");
            int gold = int.Parse(Console.ReadLine());

            Console.Write("Серебряные медали (В): ");
            int silver = int.Parse(Console.ReadLine());

            Console.Write("Бронзовые медали (К): ");
            int bronze = int.Parse(Console.ReadLine());

            Console.Write("Победы в нац. кубке (ФК): ");
            int ncWins = int.Parse(Console.ReadLine());

            Console.Write("Поражения в нац. кубке (Ф): ");
            int ncLosses = int.Parse(Console.ReadLine());

            Console.Write("Победы в ЛЧ: ");
            int clWins = int.Parse(Console.ReadLine());

            Console.Write("Поражения в ЛЧ (ФЛЧ): ");
            int clLosses = int.Parse(Console.ReadLine());

            Console.Write("Победы в ЛЕ: ");
            int elWins = int.Parse(Console.ReadLine());

            Console.Write("Поражения в ЛЕ (ФЛЕ): ");
            int elLosses = int.Parse(Console.ReadLine());

            Console.Write("Победы в КОК: ");
            int cwcWins = int.Parse(Console.ReadLine());

            Console.Write("Поражения в КОК (ФКОК): ");
            int cwcLosses = int.Parse(Console.ReadLine());

            Console.Write("Победы в ЛК: ");
            int confWins = int.Parse(Console.ReadLine());

            Console.Write("Поражения в ЛК (ФЛК): ");
            int confLosses = int.Parse(Console.ReadLine());

            var achievement = new Achievement(
                clubId, gold, silver, bronze, ncWins, ncLosses,
                clWins, clLosses, elWins, elLosses, cwcWins, cwcLosses,
                confWins, confLosses);

            dbHelper.AddAchievement(achievement);
        }

        static void QueryMenu()
        {
            Console.WriteLine("1. Клубы, выигравшие Лигу Чемпионов (1 таблица, перечень)");
            Console.WriteLine("2. Общее количество трофеев для клуба (2 таблицы, одно значение)");
            Console.WriteLine("3. Клубы с их странами и трофеями (3 таблицы, перечень)");
            Console.WriteLine("4. Страна с наибольшим количеством трофеев (3 таблицы, одно значение)");
            Console.WriteLine("5. Клубы, выигравшие еврокубки (3 таблицы, перечень)");
            Console.WriteLine("6. Клуб с наибольшим количеством серебра (3 таблицы, одно значение)");
            Console.Write("Выберите запрос: ");

            string choice = Console.ReadLine();

            switch (choice)
            {
                case "1":
                    dbHelper.GetChampionsLeagueWinners();
                    break;

                case "2":
                    Console.Write("Введите название клуба: ");
                    string clubName = Console.ReadLine();
                    dbHelper.GetTotalTrophiesByClubName(clubName);
                    break;

                case "3":
                    dbHelper.GetClubsWithCountriesAndTotalTrophies();
                    break;

                case "4":
                    dbHelper.GetCountryWithMostTrophies();
                    break;

                case "5":
                    dbHelper.GetClubsWithEuropeanTrophies();
                    break;

                case "6":
                    dbHelper.GetClubWithMostSilverMedals();
                    break;

                default:
                    Console.WriteLine("Неверный выбор!");
                    break;
            }

            Console.WriteLine("\nНажмите любую клавишу...");
            Console.ReadKey();
        }

        static void SaveData()
        {
            Console.WriteLine("\nСохранение данных");
            dbHelper.SaveToExcel();
            Console.WriteLine("\nНажмите любую клавишу...");
            Console.ReadKey();
        }
    }
}