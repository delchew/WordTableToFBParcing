using FirebirdDatabaseProvider.Attributes;

namespace CableDataParsing.TableEntityes
{
    [FBTableName("LIST_CABLE_POWER_COLOR")]
    public class ListCablePowerColor
    {
        [FBTableField(TableFieldName = "LIST_ID", TypeName = "INTEGER", IsNotNull = true, IsPrymaryKey = true)]
        [FBFieldAutoincrement(GeneratorName = "LIST_PWR_CLR_ID_GEN")]
        public long ListId { get; set; }

        [FBTableField(TableFieldName = "CABLE_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long CableId { get; set; }

        [FBTableField(TableFieldName = "POWER_COLOR_SCHEME_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long PowerColorSchemeId { get; set; }
    }
}
