using FirebirdDatabaseProvider.Attributes;

namespace CableDataParsing.TableEntityes
{
    [FBTableName("LIST_CABLE_PROPERTIES")]
    public class ListCableProperties
    {
        [FBTableField(TableFieldName = "ID", TypeName = "INTEGER", IsNotNull = true, IsPrymaryKey = true)]
        [FBFieldAutoincrement(GeneratorName = "LIST_CABLE_PROPERTIES_ID_GEN")]
        public long ListId { get; set; }

        [FBTableField(TableFieldName = "CABLE_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long CableId { get; set; }

        [FBTableField(TableFieldName = "PROPERTY_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long PropertyId { get; set; }
    }
}
