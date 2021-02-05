using GetInfoFromWordToFireBirdTable.Attributes;

namespace GetInfoFromWordToFireBirdTable.CableEntityes
{
    [FBTableName("INSULATED_BILLET")]
    public class CableBillet
    {
        [FBTableField(TableFieldName = "INS_BILLET_ID", TypeName = "INTEGER", IsNotNull = true, IsPrymaryKey = true)]
        [FBFieldAutoincrement(GeneratorsName = "INS_BILLET_GEN")]
        public long BilletId { get; set; }

        [FBTableField(TableFieldName = "COND_ID", TypeName = "INTEGER", IsNotNull = true)]
        public int ConductorId { get; set; }

        [FBTableField(TableFieldName = "POLYMER_GROUP_ID", TypeName = "INTEGER", IsNotNull = true)]
        public int PolymerGroupId { get; set; }

        [FBTableField(TableFieldName = "DIAMETER", TypeName = "NUMERIC(3, 2)", IsNotNull = true)]
        public double Diameter { get; set; }

        [FBTableField(TableFieldName = "MIN_THICKNESS", TypeName = "NUMERIC(2, 2)", IsNotNull = true)]
        public double MinThickness { get; set; }

        [FBTableField(TableFieldName = "NOMINAL_THICKNESS", TypeName = "NUMERIC(2, 2)")]
        public double? NominalThickness { get; set; }

        [FBTableField(TableFieldName = "SHRT_N_ID", TypeName = "INTEGER", IsNotNull = true)]
        public int CableShortNameId { get; set; }
    }
}
