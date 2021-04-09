using WordObj = Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.IO;
using System;

namespace CableDataParsing.MSWordTableParsers
{
    public class MSWordTableParser : WordTableParser<TableCellData>
    {
        private readonly WordObj.Application _app;

        /// <summary>
        /// Количество таблиц в документе. Выбросит исключение, если документ не был открыт до обращения к свойству
        /// </summary>
        public override int DocumentTablesCount => _app.ActiveDocument.Tables.Count;

        public MSWordTableParser(TableParserConfigurator configurator) : base(configurator)
        { }

        public MSWordTableParser() : base()
        {
            _app = new WordObj.Application { Visible = false };
        }

        /// <summary>
        /// Вызывается для открытия документа прежде чем выполнить метод парсинга
        /// </summary>
        /// <param name="mSWordFile">Файл Microsoft Word для открытия</param>
        public override void OpenWordDocument(FileInfo mSWordFile)
        {
            if (!mSWordFile.Exists)
                throw new FileNotFoundException("указанный файл отсутствует!");
            object fileName = mSWordFile.FullName;
            _app.Documents.Open(ref fileName);
        }

        /// <summary>
        /// Закрывает открытое приложение Microsoft Word
        /// </summary>
        public override void CloseWordApp()
        {
            _app.Quit();
        }

        /// <summary>
        /// Возвращает коллекцию разобраных ячеек таблицы
        /// </summary>
        /// <param name="tableNumber">Номер таблицы документа MSWord. Внимание! В Microsoft Word таблицы нумеруются с 1, а не с 0!</param>
        /// <returns></returns>
        public override IEnumerable<TableCellData> GetCableCellsCollection(int tableNumber)
        {
            try
            {
                var tables = _app.ActiveDocument.Tables;
                if (tables.Count == 0)
                    throw new Exception("Отсутствуют таблицы для парсинга в указанном Word файле!");
                var table = tables[tableNumber];
                var tableCellDataList = new List<TableCellData>();
                for (int i = 0; i < Configurator.DataRowsCount; i++)
                {
                    for (int j = 0; j < Configurator.DataColumnsCount; j++)
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
                RowHeaderData = GetStringFromWordTableCell(table, Configurator.DataStartRowIndex + rowIndex, Configurator.RowHeadersColumnIndex),
                ColumnHeaderData = GetStringFromWordTableCell(table, Configurator.ColumnHeadersRowIndex, Configurator.DataStartColumnIndex + columnIndex),
                CellData = GetStringFromWordTableCell(table, Configurator.DataStartRowIndex + rowIndex, Configurator.DataStartColumnIndex + columnIndex)
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
