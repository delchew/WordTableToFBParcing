namespace CableDataParsing.MSWordTableParsers
{
    public class TableParserConfigurator
    {
        public int DataStartRowIndex { get; set; }
        public int DataStartColumnIndex { get; set; }
        public int DataColumnsCount { get; set; }
        public int DataRowsCount { get; set; }
        public int ColumnHeadersRowIndex { get; set; }
        public int RowHeadersColumnIndex { get; set; }

        public TableParserConfigurator()
        {}

        public TableParserConfigurator(int dataStartRowIndex, int dataStartColumnIndex, int dataColumnsCount, int dataRowsCount, int columnHeadersRowIndex, int rowHeadersColumnIndex)
        {
            DataStartRowIndex = dataStartRowIndex;
            DataStartColumnIndex = dataStartColumnIndex;
            DataColumnsCount = dataColumnsCount;
            DataRowsCount = dataRowsCount;
            ColumnHeadersRowIndex = columnHeadersRowIndex;
            RowHeadersColumnIndex = rowHeadersColumnIndex;
        }

        public TableParserConfigurator Clone()
        {
            return new TableParserConfigurator
            {
                DataStartRowIndex = this.DataStartRowIndex,
                DataStartColumnIndex = this.DataStartColumnIndex,
                DataColumnsCount = this.DataColumnsCount,
                DataRowsCount = this.DataRowsCount,
                ColumnHeadersRowIndex = this.ColumnHeadersRowIndex,
                RowHeadersColumnIndex = this.RowHeadersColumnIndex
            };
        }
        public TableParserConfigurator SetDataStartRowIndex(int index)
        {
            DataStartRowIndex = index;
            return this;
        }

        public TableParserConfigurator SetDataStartColumnIndex(int index)
        {
            DataStartColumnIndex = index;
            return this;
        }

        public TableParserConfigurator SetDataRowsCount(int rowsCount)
        {
            DataRowsCount = rowsCount;
            return this;
        }

        public TableParserConfigurator SetDataColumnsCount(int columnsCount)
        {
            DataColumnsCount = columnsCount;
            return this;
        }

        public TableParserConfigurator SetColumnHeadersRowIndex(int index)
        {
            ColumnHeadersRowIndex = index;
            return this;
        }

        public TableParserConfigurator SetRowHeadersColumnIndex(int index)
        {
            RowHeadersColumnIndex = index;
            return this;
        }
    }
}
