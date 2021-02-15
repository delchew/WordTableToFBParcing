using FirebirdDatabaseProvider.Attributes;

namespace CableDataParsing.TableEntityes
{
    [FBTableName("CONDUCTOR")]
    public class ConductorPresenter
    {
        [FBTableField(TableFieldName = "ID", TypeName = "INTEGER", IsNotNull = true, IsPrymaryKey = true)]
        [FBFieldAutoincrement(GeneratorName = "CONDUCTOR_ID_GEN")]
        public long ConductorId { get; set; }

        [FBTableField(TableFieldName = "TITLE", TypeName = "VARCHAR(20)", IsNotNull = true)]
        public string Title { get; set; }

        [FBTableField(TableFieldName = "WIRES_COUNT", TypeName = "INTEGER", IsNotNull = true)]
        public int WiresCount { get; set; }

        [FBTableField(TableFieldName = "METAL_ID", TypeName = "INTEGER", IsNotNull = true)]
        public long MetalId { get; set; }

        [FBTableField(TableFieldName = "CLASS", TypeName = "INTEGER", IsNotNull = true)]
        public int Class { get; set; }

        [FBTableField(TableFieldName = "WIRES_DIAMETER", TypeName = "NUMERIC(3, 3)", IsNotNull = true)]
        public decimal WiresDiameter { get; set; }

        [FBTableField(TableFieldName = "CONDUCTOR_DIAMETER", TypeName = "NUMERIC(3, 3)", IsNotNull = true)]
        public decimal ConductorDiameter { get; set; }

        [FBTableField(TableFieldName = "AREA_IN_SQR_MM", TypeName = "NUMERIC(3, 3)", IsNotNull = true)]
        public decimal AreaInSqrMm { get; set; }
    }
}