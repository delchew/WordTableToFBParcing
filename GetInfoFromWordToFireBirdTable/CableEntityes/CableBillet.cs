using FirebirdDatabaseProvider.Attributes;

namespace GetInfoFromWordToFireBirdTable.CableEntityes
{
    [FBTableName("INSULATED_BILLET")]
    public class CableBillet
    {
        [FBTableField(TableFieldName = "INS_BILLET_ID", TypeName = "INTEGER", IsNotNull = true, IsPrymaryKey = true)]
        [FBFieldAutoincrement(GeneratorName = "INS_BILLET_GEN")]
        public long BilletId { get; set; }

        [FBTableField(TableFieldName = "COND_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long ConductorId { get; set; }

        [FBTableField(TableFieldName = "POLYMER_GROUP_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long PolymerGroupId { get; set; }

        [FBTableField(TableFieldName = "DIAMETER", TypeName = "NUMERIC(3, 2)", IsNotNull = true)]
        public decimal Diameter { get; set; }

        [FBTableField(TableFieldName = "MIN_THICKNESS", TypeName = "NUMERIC(2, 2)", IsNotNull = true)]
        public decimal MinThickness { get; set; }

        [FBTableField(TableFieldName = "NOMINAL_THICKNESS", TypeName = "NUMERIC(2, 2)")]
        public decimal? NominalThickness { get; set; }

        [FBTableField(TableFieldName = "SHRT_N_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long CableShortNameId { get; set; }
    }
}
