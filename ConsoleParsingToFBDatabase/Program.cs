﻿using GetInfoFromWordToFireBirdTable;
using GetInfoFromWordToFireBirdTable.CableEntityes;
using GetInfoFromWordToFireBirdTable.Common;
using System;
using System.IO;

namespace ConsoleParsingToFBDatabase
{
    class Program
    {
        static void Main()
        {
            ParseSkabBillet();
        }

        static void ParseSkabBillet()
        {
            var dbFile = new FileInfo(@"E:\databases\database_repository\CABLES_BRANDS.FDB");
            var parser = new SkabInsulatedBilletParser(dbFile);
            int recordsCount = 0;
            Console.WriteLine("Нажмите любую клавишу для начала парсинга...");
            Console.ReadKey();
            try
            {
                Console.WriteLine("Тест");
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

        private void Select()
        {
            var dbFile = new FileInfo(@"E:\databases\database_repository\CABLES_BRANDS.FDB");
            var parser = new SkabParser(dbFile, null);
            //var billetId1 = parser.GetInsBilletId(4, 660, 0.5);
            //var billetId2 = parser.GetInsBilletId(4, 660, 1.5);
            EndOfProgram();
        }

        private static void CreateTable()
        {
            //Users/Shared/databases/hello.fdb

            var dbFile = new FileInfo(@"E:\databases\CABLES_BRANDS.FDB");
            //var dbProvider = new FirebirdDBTableProvider<Kunrs>(dbFile);
            var dbProvider = new FirebirdDBTableProvider<Skab>(dbFile);
            //FirebirdDBTableProvider<Skab>.CreateBoolIntDBDomain(dbFile, "BOOLEAN_INT");
            dbProvider.OpenConnection();
            dbProvider.CreateTableIfNotExists();
            dbProvider.CloseConnection();
        }

        private static void EndOfProgram()
        {
            Console.Write(Environment.NewLine);
            Console.WriteLine("Программа завершена. Нажмите любую клавишу...");
            Console.ReadKey();
        }
    }
}
