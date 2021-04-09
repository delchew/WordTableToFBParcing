using System.Collections.Generic;
using System.IO;

namespace CableDataParsing.MSWordTableParsers
{
    public abstract class WordTableParser<T> where T : new()
    {
        public TableParserConfigurator Configurator { get; set; }

        public abstract int DocumentTablesCount { get; }

        public WordTableParser()
        {
            Configurator = new TableParserConfigurator();
        }

        public WordTableParser(TableParserConfigurator configurator)
        {
            Configurator = configurator;
        }
        
        
        public IEnumerable<T> GetCableCellsCollection(int tableNumber, TableParserConfigurator configurator)
        {
            Configurator = configurator;
            return GetCableCellsCollection(tableNumber);
        }

        public abstract IEnumerable<T> GetCableCellsCollection(int tableNumber);

        public abstract void OpenWordDocument(FileInfo _mSWordFile);

        /// <summary>
        /// Вызывается для открытия документа прежде чем выполнить метод парсинга
        /// </summary>
        /// <param name="msWordFilePath">Полный путь к файлу Microsoft Word для открытия</param>
        public void OpenWordDocument(string msWordFilePath)
        {
            OpenWordDocument(new FileInfo(msWordFilePath));
        }

        public abstract void CloseWordApp();
    }
}
