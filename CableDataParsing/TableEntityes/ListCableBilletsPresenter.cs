using FirebirdDatabaseProvider.Attributes;

namespace CableDataParsing.TableEntityes
{
    [FBTableName("LIST_CABLE_BILLETS")]
    public class ListCableBilletsPresenter
    {
        [FBTableField(TableFieldName = "ID", TypeName = "INTEGER", IsNotNull = true, IsPrymaryKey = true)]
        [FBFieldAutoincrement(GeneratorName = "LIST_CABLE_BILLETS_ID_GEN")]
        public long Id { get; set; }

        [FBTableField(TableFieldName = "CABLE_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long CableId { get; set; }

        [FBTableField(TableFieldName = "BILLET_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long BilletId { get; set; }
    }
}
