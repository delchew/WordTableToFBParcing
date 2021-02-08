using GetInfoFromWordToFireBirdTable;
using GetInfoFromWordToFireBirdTable.CableEntityes;
using GetInfoFromWordToFireBirdTable.Common;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace ConsoleParsingToFBDatabase
{
    class Program
    {
        static void Main()
        {
            CreateTableAndParsingsSkabData();
        }

        static void CreateTableAndParsingsSkabData()
        {
            //var provider = new FirebirdDBTableProvider<Skab>();
            //provider.OpenConnection();
            //provider.CreateTableIfNotExists();
            //provider.CloseConnection();

            var parser = new SkabParser(null);
            //parser.ParseReport += (num1, num2) => { Console.Write($"{num2}.."); };
            //var result = await Task<int>.Factory.StartNew(parser.ParseDataToDatabase);
            var result = parser.ParseDataToDatabase();
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Добавлено {0} записей.", result);
            Console.ReadKey();
        }

        static void MakeSelect()
        {
            var provider = new FirebirdDBTableProvider<CableBillet>();
            try
            {
                provider.OpenConnection();
                if (provider.TableExists())
                {
                    var result = provider.GetAllItemsFromTable();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Произошла ошибка: {0}", ex.Message);
            }
            finally
            {
                provider.CloseConnection();
                EndOfProgram();
            }
        }

        private static void ParseComplete(int result)
        {
            Console.WriteLine("Добавлено {0} записей.", result);
        }

        static void CreateTableAndParsingBilletData()
        {
            int recordsCount = 0;
            try
            {
                var provider = new FirebirdDBTableProvider<CableBillet>();
                provider.OpenConnection();
                provider.CreateTableIfNotExists();
                provider.CloseConnection();
                var parser = new SkabInsulatedBilletParser();
                recordsCount = parser.ParseDataToDatabase();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Произошла ошибка: {0}", ex.Message);
            }
            finally
            {
                Console.WriteLine("Добавлено {0} записей.", recordsCount);
                EndOfProgram();
            }
        }

        private static void ParseKunrsCable()
        {
            var fileInfo = new FileInfo(@"C:\Users\a.bondarenko\Desktop\kunrs.docx");
            var parser = new KunrsParser(fileInfo);
            int recordsCount = 0;
            Console.WriteLine("Нажмите любую клавишу для начала парсинга...");
            Console.ReadKey();
            try
            {
                //Console.WriteLine("Тест");
                recordsCount = parser.ParseDataToDatabase();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Произошла ошибка: {0}", ex.Message);
            }
            finally
            {
                Console.WriteLine("Добавлено {0} записей.", recordsCount);
                EndOfProgram();
            }
        }
        private static void ParseSkabBillet()
        {
            var parser = new SkabInsulatedBilletParser();
            int recordsCount = 0;
            //Console.WriteLine("Нажмите любую клавишу для начала парсинга...");
            //Console.ReadKey();
            try
            {
                Console.WriteLine("Тест");
                recordsCount = parser.ParseDataToDatabase();
                Console.WriteLine("Добавлено {0} записей.", recordsCount);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Произошла ошибка: {0}", ex.Message);
            }
            finally
            {
                EndOfProgram();
            }
        }

        private static void Select()
        {
            var parser = new SkabParser(null);
            //var billetId1 = parser.GetInsBilletId(4, 660, 0.5);
            //var billetId2 = parser.GetInsBilletId(4, 660, 1.5);
            EndOfProgram();
        }

        private static void CreateKunrsTable()
        {
            var dbProvider = new FirebirdDBTableProvider<Kunrs>();
            dbProvider.OpenConnection();
            try
            {
                //dbProvider.CreateBoolIntDBDomain("BOOLEAN_INT");
                dbProvider.CreateTableIfNotExists();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Произошла ошибка: {0}", ex.Message);
            }
            finally
            {
                dbProvider.CloseConnection();
                EndOfProgram();
            }

        }

        private static void EndOfProgram()
        {
            Console.Write(Environment.NewLine);
            Console.WriteLine("Программа завершена. Нажмите любую клавишу...");
            Console.ReadKey();
        }
    }
}
