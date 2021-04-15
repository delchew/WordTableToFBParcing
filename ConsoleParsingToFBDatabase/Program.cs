using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using CableDataParsing;
using CableDataParsing.TableEntityes;
using Cables;
using CablesCraftMobile;
using FirebirdDatabaseProvider;
using FirebirdSql.Data.FirebirdClient;

namespace ConsoleParsingToFBDatabase
{
    class Program
    {
        static readonly string _connectionString = "character set=utf8;user id=SYSDBA;password=masterkey;dialect=3;data source=localhost;port number=3050;initial catalog=/Users/Shared/databases/database_repository/CablesDatabases/CABLES.FDB";
        static readonly string _connectionString2 = "character set=utf8;user id=SYSDBA;password=masterkey;dialect=3;data source=localhost;port number=3050;initial catalog=e:\\databases\\CABLES.FDB";
        static readonly string _connectionString3 = "character set=utf8;user id=SYSDBA;password=masterkey;dialect=3;data source=localhost;port number=3050;initial catalog=e:\\databases\\CABLESBRANDS.FDB";
        static readonly string _connectionString4 = "character set=utf8;user id=SYSDBA;password=masterkey;dialect=3;data source=localhost;port number=3050;initial catalog=e:\\databases\\database_repository\\CABLES.FDB";
        static readonly string _connectionString5 = "character set=utf8;user id=SYSDBA;password=masterkey;dialect=3;data source=localhost;port number=3050;initial catalog=/Users/Shared/databases/CABLES1.FDB";
        static readonly string _wordFilePath = @"E:\CableDiametersTables\Kpsvev.docx";
        static readonly string _wordFilePath1 = @"/Users/Shared/databases/database_repository/CablesDatabases/CableDiametersTables/Kpsvev.docx";
        static readonly string _wordFilePath2 = @"/Users/Shared/databases/database_repository/CablesDatabases/CableDiametersTables/Skab.docx";

        static void Main()
        {
            using var parser = new KpsvevParser(_connectionString5, new FileInfo(_wordFilePath1));
            //using var parser = new SkabParser(_connectionString5, new FileInfo(_wordFilePath2));
            parser.ParseReport += Parser_ParseReport;
            var recordsCount = parser.ParseDataToDatabase();
            Console.WriteLine("{0} записей внесено в базу.", recordsCount);

            Console.ReadKey();
        }

        private static void Parser_ParseReport(double percentage)
        {
            Console.Clear();
            Console.WriteLine(percentage * 100 + "% выполнено...");
        }

        static void WorkingWithADONetArrays()
        {
            using (var connection = new FbConnection(_connectionString3))
            {
                connection.Open();
                var sqlSelectQueryString = $"SELECT * FROM TWIST_INFO;";
                FbDataReader reader;
                using (var command = new FbCommand(sqlSelectQueryString, connection))
                {
                    reader = command.ExecuteReader();
                    if (reader.HasRows)
                    {
                        while (reader.Read())
                        {
                            var values = new object[reader.FieldCount];
                            reader.GetValues(values);
                            Console.WriteLine(string.Join("|", values));

                            for (int i = 0; i < reader.FieldCount; i++)
                            {
                                var obj = reader.GetValue(i);
                                var type = obj.GetType();
                                if (obj is int[] arr)
                                {
                                    for (int j = 0; j < arr.Length; j++)
                                        Console.Write($"{arr[j]}, ");
                                }
                                else
                                    Console.Write($"{obj} ");
                            }
                            Console.WriteLine();
                        }
                    }
                }
                sqlSelectQueryString = "INSERT INTO TWIST_INFO (ELEMENTS_COUNT, TWIST_COEFFICIENT, LAYERS_ELEMENTS_COUNT) VALUES (:V1, :v2, :V3);";
                using (var command = new FbCommand(sqlSelectQueryString, connection))
                {
                    (int cnt, double koef, int[] lyrs)[] paramSet = new (int cnt, double koef, int[] lyrs)[]
                    {
                        (49, 8.41, new int []{ 4, 9, 15, 21, 0} ),
                        (50, 8.41, new int []{ 4, 10, 15, 21, 0} ),
                        (61, 9.0, new int []{ 1, 6, 12, 18, 24} ),
                        (20, 5.3, new int []{ 7, 13, 0, 0, 0} ),
                    };
                    foreach (var set in paramSet)
                    {
                        command.Parameters.Add(new FbParameter("V1", set.cnt));
                        command.Parameters.Add(new FbParameter("V2", set.koef));
                        command.Parameters.Add("V3", set.lyrs);
                        command.ExecuteNonQuery();
                        command.Parameters.Clear();
                    }
                }

            }
        }

        static void FillTwistInfoTable()
        {
            var twistInfoList = GetTwistInfoList();
            Console.WriteLine("Число записей: {0}", twistInfoList.Count);
            var dbProvider = new FirebirdDBProvider(_connectionString2);
            var provider = new FirebirdDBTableProvider<TwistInfoPresenter>(dbProvider);
            var presenter = new TwistInfoPresenter();

            try
            {
                dbProvider.OpenConnection();


                if (provider.TableExists())
                {
                    PropertyInfo prop;
                    foreach (var twistInfo in twistInfoList)
                    {
                        presenter.ElementsCount = twistInfo.QuantityElements;
                        presenter.TwistKoefficient = (decimal)twistInfo.TwistCoefficient;
                        for (int i = 0; i < twistInfo.LayersElementsCount.Length; i++)
                        {
                            prop = presenter.GetType().GetProperty($"Layer{i + 1}ElementsCount");
                            prop.SetValue(presenter, twistInfo.LayersElementsCount[i]);
                        }
                        provider.AddItem(presenter);
                    }
                }
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
            var parser = new BilletParser(_connectionString);
            int recordsCount = 0;
            Console.WriteLine("Нажмите любую клавишу для начала парсинга...");
            Console.ReadKey();
            try
            {
                //recordsCount = parser.ParseDataToDatabase();
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
