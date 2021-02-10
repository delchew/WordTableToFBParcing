using FirebirdDatabaseProvider;
using GetInfoFromWordToFireBirdTable.TableEntityes;
using GetInfoFromWordToFireBirdTable.Common;
using System;
using System.IO;

namespace ConsoleParsingToFBDatabase
{
    class Program
    {
        static readonly string _connectionString = "character set=utf8;user id=SYSDBA;password=masterkey;dialect=3;data source=localhost;port number=3050;initial catalog=E:\\databases\\TEST.fdb";

        static void Main()
        {
            var kunrs = new KunrsPresenter
            {
                
            };
        }

        static void CreateConductorTable()
        {
            var provider = new FirebirdDBTableProvider<ConductorPresenter>(_connectionString);
            provider.OpenConnection();
            provider.CreateTableIfNotExists();
            provider.CloseConnection();
        }

        static void MakeSelect()
        {
            var provider = new FirebirdDBTableProvider<CableBilletPresenter>(_connectionString);
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

        static void SkabConductorsParse()
        {
            var parser = new SkabConductorsParcer(_connectionString);
            Console.WriteLine("Нажмите любую клавишу для начала парсинга...");
            Console.ReadKey();
            int recordsCount = 0;
            try
            {
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

        private static void ParseComplete(int result)
        {
            Console.WriteLine("Добавлено {0} записей.", result);
        }

        static void CreateTableAndParsingBilletData()
        {
            int recordsCount = 0;
            try
            {
                //var provider = new FirebirdDBTableProvider<CableBillet>(connectionString);
                //provider.OpenConnection();
                //provider.CreateTableIfNotExists();
                //provider.CloseConnection();
                var parser = new SkabInsulatedBilletParser(_connectionString);
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
            var parser = new KunrsParser(_connectionString, fileInfo);
            int recordsCount = 0;
            Console.WriteLine("Нажмите любую клавишу для начала парсинга...");
            Console.ReadKey();
            try
            {
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
            var parser = new SkabInsulatedBilletParser(_connectionString);
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

        private static void CreateKunrsTable()
        {
            var dbProvider = new FirebirdDBTableProvider<KunrsPresenter>(_connectionString);
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
