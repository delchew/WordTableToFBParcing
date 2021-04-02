using System.Collections.Generic;
using System.IO;

namespace CableDataParsing.MSWordTableParsers
{
    public abstract class WordTableParser<T> where T : new()
    {
        public int DataStartRowIndex { get; set; }
        public int DataStartColumnIndex { get; set; }
        public int DataColumnsCount { get; set; }
        public int DataRowsCount { get; set; }
        public int ColumnHeadersRowIndex { get; set; }
        public int RowHeadersColumnIndex { get; set; }

        public WordTableParser<T> SetDataStartRowIndex(int index)
        {
            DataStartRowIndex = index;
            return this;
        }

        public WordTableParser<T> SetDataStartColumnIndex(int index)
        {
            DataStartColumnIndex = index;
            return this;
        }

        public WordTableParser<T> SetDataRowsCount(int rowsCount)
        {
            DataRowsCount = rowsCount;
            return this;
        }

        public WordTableParser<T> SetDataColumnsCount(int columnsCount)
        {
            DataColumnsCount = columnsCount;
            return this;
        }

        public WordTableParser<T> SetColumnHeadersRowIndex(int index)
        {
            ColumnHeadersRowIndex = index;
            return this;
        }

        public WordTableParser<T> SetRowHeadersColumnIndex(int index)
        {
            RowHeadersColumnIndex = index;
            return this;
        }
        public abstract List<T> GetCableCellsCollection(int tableNumber);

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
