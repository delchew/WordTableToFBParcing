using FirebirdDatabaseProvider.Attributes;

namespace CableDataParsing.TableEntityes
{
    [FBTableName("LIST_CABLE_PROPERTIES")]
    public class ListCableProperties
    {
        [FBTableField(TableFieldName = "LIST_ID", TypeName = "INTEGER", IsNotNull = true, IsPrymaryKey = true)]
        [FBFieldAutoincrement(GeneratorName = "LIST_PROP_ID_GEN")]
        public long ListId { get; set; }

        [FBTableField(TableFieldName = "CABLE_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long CableId { get; set; }

        [FBTableField(TableFieldName = "PROP_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long PropertyId { get; set; }
    }
}
