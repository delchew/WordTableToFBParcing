using FirebirdDatabaseProvider.Attributes;

namespace CableDataParsing.TableEntityes
{
    [FBTableName("CABLE_SKAB")]
    public class SkabPresenter : CablePresenter
    {
        [FBTableField(TableFieldName = "HAS_INDIVIDUAL_FOIL_SHIELDS", TypeName = "BOOLEAN_INT", IsNotNull = true)]
        public bool HasIndividualFoilShields { get; set; }

        [FBTableField(TableFieldName = "HAS_FOIL_SHIELD", TypeName = "BOOLEAN_INT", IsNotNull = true)]
        public bool HasFoilShield { get; set; }

        [FBTableField(TableFieldName = "HAS_BRAID_SHIELD", TypeName = "BOOLEAN_INT", IsNotNull = true)]
        public bool HasBraidShield { get; set; }

        [FBTableField(TableFieldName = "HAS_FILLING", TypeName = "BOOLEAN_INT", IsNotNull = true)]
        public bool HasFilling { get; set; }

        [FBTableField(TableFieldName = "HAS_ARMOUR_BRAID", TypeName = "BOOLEAN_INT", IsNotNull = true)]
        public bool HasArmourBraid { get; set; }

        [FBTableField(TableFieldName = "HAS_ARMOUR_TUBE", TypeName = "BOOLEAN_INT", IsNotNull = true)]
        public bool HasArmourTube { get; set; }

        [FBTableField(TableFieldName = "HAS_WATERBLOCK_STRIPE", TypeName = "BOOLEAN_INT", IsNotNull = true)]
        public bool HasWaterBlockStripe { get; set; }

        [FBTableField(TableFieldName = "SPARK_SAFETY", TypeName = "BOOLEAN_INT", IsNotNull = true)]
        public bool SparkSafety { get; set; }

    }
}