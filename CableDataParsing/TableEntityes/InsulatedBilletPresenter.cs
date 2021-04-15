using FirebirdDatabaseProvider.Attributes;

namespace CableDataParsing.TableEntityes
{
    [FBTableName("INSULATED_BILLET")]
    public class InsulatedBilletPresenter
    {
        [FBTableField(TableFieldName = "ID", TypeName = "INTEGER", IsNotNull = true, IsPrymaryKey = true)]
        [FBFieldAutoincrement(GeneratorName = "INSULATED_BILLET_ID_GEN")]
        public long Id { get; set; }

        [FBTableField(TableFieldName = "COND_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long ConductorId { get; set; }

        [FBTableField(TableFieldName = "POLYMER_GROUP_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long PolymerGroupId { get; set; }

        [FBTableField(TableFieldName = "OPERATING_VOLTAGE_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long OperatingVoltageId { get; set; }

        [FBTableField(TableFieldName = "DIAMETER", TypeName = "NUMERIC(3, 2)", IsNotNull = true)]
        public decimal Diameter { get; set; }

        [FBTableField(TableFieldName = "MIN_THICKNESS", TypeName = "NUMERIC(2, 2)", IsNotNull = true)]
        public decimal? MinThickness { get; set; }

        [FBTableField(TableFieldName = "NOMINAL_THICKNESS", TypeName = "NUMERIC(2, 2)", IsNotNull = true)]
        public decimal? NominalThickness { get; set; }

        [FBTableField(TableFieldName = "CABLE_SHORT_NAME_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long CableShortNameId { get; set; }
    }
}