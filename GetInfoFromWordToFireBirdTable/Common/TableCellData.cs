namespace GetInfoFromWordToFireBirdTable.Common
{
    public class TableCellData
    {
        public string ColumnHeaderData { get; set; }
        public string RowHeaderData { get; set; }
        public string CellData { get; set; }

        public override string ToString()
        {
            return $"{ColumnHeaderData} х {RowHeaderData} - {CellData}";
        }
    }
}
