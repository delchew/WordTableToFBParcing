using FirebirdDatabaseProvider.Attributes;

namespace GetInfoFromWordToFireBirdTable.CableEntityes
{
    [FBTableName("CABLE_KUNRS")]
    public class Kunrs : Cable
    {
        [FBTableField(TableFieldName = "HAS_FOIL_SHIELD", TypeName = "BOOLEAN_INT", IsNotNull = true)]
        public bool HasFoilShield { get; set; }

        [FBTableField(TableFieldName = "HAS_FILLING", TypeName = "BOOLEAN_INT", IsNotNull = true)]
        public bool HasFilling { get; set; }

        [FBTableField(TableFieldName = "HAS_ARMOUR_BRAID", TypeName = "BOOLEAN_INT", IsNotNull = true)]
        public bool HasArmourBraid { get; set; }

        [FBTableField(TableFieldName = "HAS_ARMOUR_TUBE", TypeName = "BOOLEAN_INT", IsNotNull = true)]
        public bool HasArmourTube { get; set; }

        [FBTableField(TableFieldName = "PWR_COLOR_SCHEME_ID", TypeName = "INTEGER", IsNotNull = true)]
        public int PowerColorSchemeId { get; set; }
    }
}
