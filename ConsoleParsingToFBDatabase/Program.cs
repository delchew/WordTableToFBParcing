using GetInfoFromWordToFireBirdTable;
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
            CreateTable();
            //ParseKunrsCable();
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

        private static void Select()
        {
            var parser = new SkabParser(null);
            //var billetId1 = parser.GetInsBilletId(4, 660, 0.5);
            //var billetId2 = parser.GetInsBilletId(4, 660, 1.5);
            EndOfProgram();
        }

        private static void CreateTable()
        {
            //var dbProvider = new FirebirdDBTableProvider<Kunrs>(dbFile);
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
