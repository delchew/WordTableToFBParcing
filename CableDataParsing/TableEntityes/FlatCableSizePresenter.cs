using FirebirdDatabaseProvider.Attributes;

namespace CableDataParsing.TableEntityes
{
    [FBTableName("FLAT_CABLE_SIZE")]
    public class FlatCableSizePresenter
    {
        [FBTableField(TableFieldName = "ID", TypeName = "INTEGER", IsNotNull = true, IsPrymaryKey = true)]
        [FBFieldAutoincrement(GeneratorName = "FLAT_CABLE_SIZE_ID_GEN")]
        public long Id { get; set; }

        [FBTableField(TableFieldName = "HEIGHT", TypeName = "NUMERIC(3, 2)", IsNotNull = true)]
        public decimal Height { get; set; }

        [FBTableField(TableFieldName = "WIDTH", TypeName = "NUMERIC(3, 2)", IsNotNull = true)]
        public decimal Width { get; set; }

        [FBTableField(TableFieldName = "CABLE_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long CableId { get; set; }


    }
}
