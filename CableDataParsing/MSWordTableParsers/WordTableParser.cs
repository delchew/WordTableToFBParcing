using WordObj = Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.IO;
using System;

namespace CableDataParsing.MSWordTableParsers
{
    public class WordTableParser : IWordTableParser<TableCellData>
    {
        private WordObj.Application _app;

        public int DataStartRowIndex { get; set; }
        public int DataStartColumnIndex { get; set; }
        public int DataColumnsCount { get; set; }
        public int DataRowsCount { get; set; }
        public int ColumnHeadersRowIndex { get; set; }
        public int RowHeadersColumnIndex { get; set; }

        public WordTableParser()
        {
            _app = new WordObj.Application { Visible = false };
        }

        public WordTableParser SetDataStartRowIndex(int index)
        {
            DataStartRowIndex = index;
            return this;
        }

        public WordTableParser SetDataStartColumnIndex(int index)
        {
            DataStartColumnIndex = index;
            return this;
        }

        public WordTableParser SetDataRowsCount(int rowsCount)
        {
            DataRowsCount = rowsCount;
            return this;
        }

        public WordTableParser SetDataColumnsCount(int columnsCount)
        {
            DataColumnsCount = columnsCount;
            return this;
        }

        public WordTableParser SetColumnHeadersRowIndex(int index)
        {
            ColumnHeadersRowIndex = index;
            return this;
        }

        public WordTableParser SetRowHeadersColumnIndex(int index)
        {
            RowHeadersColumnIndex = index;
            return this;
        }
        /// <summary>
        /// Вызывается для открытия документа прежде чем выполнить метод парсинга
        /// </summary>
        /// <param name="mSWordFile">Файл Microsoft Word для открытия</param>
        public void OpenWordDocument(FileInfo mSWordFile)
        {
            if (!mSWordFile.Exists)
                throw new FileNotFoundException("указанный файл отсутствует!");
            object fileName = mSWordFile.FullName;
            _app.Documents.Open(ref fileName);
        }

        /// <summary>
        /// Вызывается для открытия документа прежде чем выполнить метод парсинга
        /// </summary>
        /// <param name="msWordFilePath">Полный путь к файлу Microsoft Word для открытия</param>
        public void OpenWordDocument(string msWordFilePath)
        {
            OpenWordDocument(new FileInfo(msWordFilePath));
        }

        /// <summary>
        /// Закрывает открытое приложение Microsoft Word
        /// </summary>
        public void CloseWordApp()
        {
            _app.Quit();
        }

        /// <summary>
        /// Возвращает коллекцию разобраных ячеек таблицы
        /// </summary>
        /// <param name="tableNumber">Номер таблицы документа MSWord. Внимание! В Microsoft Word таблицы нумеруются с 1, а не с 0!</param>
        /// <returns></returns>
        public List<TableCellData> GetCableCellsCollection(int tableNumber)
        {
            try
            {
                var tables = _app.ActiveDocument.Tables;
                if (tables.Count == 0)
                    throw new Exception("Отсутствуют таблицы для парсинга в указанном Word файле!");
                var table = tables[tableNumber];
                var tableCellDataList = new List<TableCellData>();
                for (int i = 0; i < DataRowsCount; i++)
                {
                    for (int j = 0; j < DataColumnsCount; j++)
                    {
                        var tableCellData = GetCellData(table, i, j);
                        tableCellDataList.Add(tableCellData);
                    }
                }
                return tableCellDataList;
            }
            catch (Exception)
            {
                CloseWordApp();
                throw;
            }
        }

        private TableCellData GetCellData(WordObj.Table table, int rowIndex, int columnIndex)
        {
            var tableCellData = new TableCellData
            {
                RowHeaderData = GetStringFromWordTableCell(table, DataStartRowIndex + rowIndex, RowHeadersColumnIndex),
                ColumnHeaderData = GetStringFromWordTableCell(table, ColumnHeadersRowIndex, DataStartColumnIndex + columnIndex),
                CellData = GetStringFromWordTableCell(table, DataStartRowIndex + rowIndex, DataStartColumnIndex + columnIndex)
            };
            return tableCellData;
        }

        private string GetStringFromWordTableCell(WordObj.Table table, int rowNumber, int columnNumber)
        {
            var cellString = table.Cell(rowNumber, columnNumber).Range.Text.Trim('\r', '\a');
            return cellString;
        }
    }
}
