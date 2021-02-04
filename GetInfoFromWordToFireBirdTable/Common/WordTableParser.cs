using WordObj = Microsoft.Office.Interop.Word;
using System.Collections.Generic;

namespace GetInfoFromWordToFireBirdTable.Common
{
    public class WordTableParser : IWordTableParser<TableCellData>
    {
        public int DataStartRowIndex { get; set; }
        public int DataStartColumnIndex { get; set; }
        public int DataColumnsCount { get; set; }
        public int DataRowsCount { get; set; }
        public int ColumnHeadersRowIndex { get; set; }
        public int RowHeadersColumnIndex { get; set; }
        public List<TableCellData> GetCableCellsCollection(WordObj.Table table)
        {
            var tableCellDataList = new List<TableCellData>();
            for (int i = 0; i < DataRowsCount; i++)
            {
                var rowHeaderData = GetStringFromWordTableCell(table, DataStartRowIndex + i, RowHeadersColumnIndex);
                for (int j = 0; j < DataColumnsCount; j++)
                {
                    var columnHeaderData = GetStringFromWordTableCell(table, ColumnHeadersRowIndex, DataStartColumnIndex + j);
                    var cellData = GetStringFromWordTableCell(table, DataStartRowIndex + i, DataStartColumnIndex + j);
                    var tableCellData = new TableCellData
                    {
                        RowHeaderData = rowHeaderData,
                        ColumnHeaderData = columnHeaderData,
                        CellData = cellData
                    };
                    tableCellDataList.Add(tableCellData);
                }
            }
            return tableCellDataList;
        }

        private string GetStringFromWordTableCell(WordObj.Table table, int rowNumber, int columnNumber)
        {
            var cellString = table.Cell(rowNumber, columnNumber).Range.Text.Trim('\r', '\a');
            return cellString;
        }
    }
}
