using GetInfoFromWordToFireBirdTable.Attributes;

namespace GetInfoFromWordToFireBirdTable.CableEntityes
{
    public abstract class Cable
    {
        [FBTableField(TableFieldName = "CABLE_ID", TypeName = "INTEGER", IsNotNull = true, IsPrymaryKey = true)]
        [FBFieldAutoincrement(GeneratorsName = "CABLE_ID_GEN")]
        public int CableId { get; set; }

        [FBTableField(TableFieldName = "BILLET_ID", TypeName = "INTEGER", IsNotNull = true)]
        public int BilletId { get; set; }

        [FBTableField(TableFieldName = "ELEMENTS_COUNT", TypeName = "INTEGER", IsNotNull = true)]
        public int ElementsCount { get; set; }

        [FBTableField(TableFieldName = "TWISTED_ELEMENT_TYPE_ID", TypeName = "INTEGER", IsNotNull = true)]
        public int TwistedElementTypeId { get; set; }

        [FBTableField(TableFieldName = "TECH_COND_ID", TypeName = "INTEGER", IsNotNull = true)]
        public int TechCondId { get; set; }

        [FBTableField(TableFieldName = "MAX_COVER_DIAMETER", TypeName = "NUMERIC(3, 2)", IsNotNull = true)]
        public double MaxCoverDiameter { get; set; }

        [FBTableField(TableFieldName = "FIRE_PROTECT_ID", TypeName = "INTEGER", IsNotNull = true)]
        public int FireProtectionId { get; set; }

        [FBTableField(TableFieldName = "COVER_POLYMER_GROUP_ID", TypeName = "INTEGER", IsNotNull = true)]
        public int CoverPolimerGroupId { get; set; }

        [FBTableField(TableFieldName = "COVER_COLOR_ID", TypeName = "INTEGER", IsNotNull = true)]
        public int CoverColorId { get; set; }
    }
}
