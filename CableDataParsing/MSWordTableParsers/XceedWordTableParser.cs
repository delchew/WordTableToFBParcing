using System;
using System.Collections.Generic;
using System.IO;
using Xceed.Document.NET;
using Xceed.Words.NET;

namespace CableDataParsing.MSWordTableParsers
{
    public class XceedWordTableParser : WordTableParser<TableCellData>
    {
        private DocX _document;

        /// <summary>
        /// Вызывается для открытия документа прежде чем выполнить метод парсинга
        /// </summary>
        /// <param name="mSWordFile">Файл Microsoft Word для открытия</param>
        public override void OpenWordDocument(FileInfo mSWordFile)
        {
            if (!mSWordFile.Exists)
                throw new FileNotFoundException("указанный файл отсутствует!");
            _document = DocX.Load(mSWordFile.FullName);
        }

        /// <summary>
        /// Закрывает открытое приложение Microsoft Word
        /// </summary>
        public override void CloseWordApp()
        {
            _document.Dispose();
        }
        public override List<TableCellData> GetCableCellsCollection(int tableNumber)
        {
            try
            {
                var tables = _document.Tables;
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

        private TableCellData GetCellData(Table table, int rowIndex, int columnIndex)
        {
            var tableCellData = new TableCellData
            {
                RowHeaderData = GetStringFromWordTableCell(table, DataStartRowIndex + rowIndex, RowHeadersColumnIndex),
                ColumnHeaderData = GetStringFromWordTableCell(table, ColumnHeadersRowIndex, DataStartColumnIndex + columnIndex),
                CellData = GetStringFromWordTableCell(table, DataStartRowIndex + rowIndex, DataStartColumnIndex + columnIndex)
            };
            return tableCellData;
        }

        private string GetStringFromWordTableCell(Table table, int rowNumber, int columnNumber)
        {
            var cellString = table.Rows[rowNumber].Cells[columnNumber].Paragraphs[0].Text.Trim('\r', '\a');
            return cellString;
        }
    }
}
