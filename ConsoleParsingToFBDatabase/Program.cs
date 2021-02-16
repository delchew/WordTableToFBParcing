using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CableDataParsing;
using CableDataParsing.TableEntityes;
using Cables;
using CablesCraftMobile;
using FirebirdDatabaseProvider;

namespace ConsoleParsingToFBDatabase
{
    class Program
    {
        static readonly string _connectionString = "character set=utf8;user id=SYSDBA;password=masterkey;dialect=3;data source=localhost;port number=3050;initial catalog=/Users/Shared/databases/database_repository/CablesDatabases/CABLES.FDB";
        static readonly string _connectionString2 = "character set=utf8;user id=SYSDBA;password=masterkey;dialect=3;data source=localhost;port number=3050;initial catalog=e:\\databases\\CABLES.FDB";
        static void Main()
        {
            var twistInfoList = GetTwistInfoList();
            Console.WriteLine("Число записей: {0}", twistInfoList.Count);
            var dbProvider = new FirebirdDBProvider(_connectionString2);
            var provider = new FirebirdDBTableProvider<TwistInfoPresenter>(dbProvider);
            var presenter = new TwistInfoPresenter();
            foreach(var twistInfo in twistInfoList)
            {
                presenter.ElementsCount = twistInfo.QuantityElements;
                presenter.TwistKoefficient = (decimal)twistInfo.TwistCoefficient;
                presenter.LayersElementsCount = twistInfo.LayersElementsCount;
                provider.AddItem(presenter);
            }
            Console.ReadKey();
        }

        static List<TwistInfo> GetTwistInfoList()
        {
            var repository = new JsonRepository();
            var filePath = Path.Combine(Environment.CurrentDirectory, "twistInfo.json");
            var list = repository.GetObjects<TwistInfo>(filePath);
            return list.ToList();
        }

        //static void CreateConductorTable()
        //{
        //    var provider = new FirebirdDBTableProvider<ConductorPresenter>(_connectionString);
        //    provider.OpenConnection();
        //    provider.CreateTableIfNotExists();
        //    provider.CloseConnection();
        //}

        //static void MakeSelect()
        //{
        //    var provider = new FirebirdDBTableProvider<CableBilletPresenter>(_connectionString);
        //    try
        //    {
        //        provider.OpenConnection();
        //        if (provider.TableExists())
        //        {
        //            var result = provider.GetAllItemsFromTable();
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Произошла ошибка: {0}", ex.Message);
        //    }
        //    finally
        //    {
        //        provider.CloseConnection();
        //        EndOfProgram();
        //    }
        //}

        static void SkabConductorsParse()
        {
            //var parser = new SkabConductorsParcer(_connectionString);
            Console.WriteLine("Нажмите любую клавишу для начала парсинга...");
            Console.ReadKey();
            int recordsCount = 0;
            try
            {
                //recordsCount = parser.ParseDataToDatabase();
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
                //var parser = new SkabInsulatedBilletParser(_connectionString);
               // recordsCount = parser.ParseDataToDatabase();
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
            //var parser = new KunrsParser(_connectionString, fileInfo);
            int recordsCount = 0;
            Console.WriteLine("Нажмите любую клавишу для начала парсинга...");
            Console.ReadKey();
            try
            {
                //recordsCount = parser.ParseDataToDatabase();
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
            Console.WriteLine("Нажмите любую клавишу для начала парсинга...");
            Console.ReadKey();
            try
            {
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

        //private static void CreateKunrsTable()
        //{
        //    var dbProvider = new FirebirdDBTableProvider<KunrsPresenter>(_connectionString);
        //    dbProvider.OpenConnection();
        //    try
        //    {
        //        //dbProvider.CreateBoolIntDBDomain("BOOLEAN_INT");
        //        dbProvider.CreateTableIfNotExists();
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine("Произошла ошибка: {0}", ex.Message);
        //    }
        //    finally
        //    {
        //        dbProvider.CloseConnection();
        //        EndOfProgram();
        //    }

        //}

        private static void EndOfProgram()
        {
            Console.Write(Environment.NewLine);
            Console.WriteLine("Программа завершена. Нажмите любую клавишу...");
            Console.ReadKey();
        }
    }
}
